#![warn(clippy::all, clippy::pedantic, clippy::nursery)]

mod consts;
mod error;
mod read;

use cfb::CompoundFile;
use clap::{Args, Parser, Subcommand};
use std::io::{prelude::*, Cursor};

use crate::error::UnlockError;
use crate::error::UnlockResult;
use crate::read::print_info;

#[derive(Parser)]
#[command(author, version, about, long_about = None)]
struct Cli {
    /// Mode to run in
    #[command(subcommand)]
    command: Commands,
}

#[derive(Subcommand)]
enum Commands {
    /// Read the contents of the Excel file
    Read(ReadArgs),

    /// Update the file to remove all protection
    Remove(RemoveArgs),
}

#[derive(Args)]
struct ReadArgs {
    /// Attempt to decode the SHA1 hash of the password
    #[arg(short, long, default_value_t = false)]
    decode: bool,

    /// Excel file to read / unlock
    filename: String,
}

#[derive(Args)]
struct RemoveArgs {
    /// Modify the file in-place, if not selected a new file will be generated and saved alongside
    /// the original
    #[arg(short, long, default_value_t = false)]
    inplace: bool,

    /// Excel file to read / unlock
    filename: String,
}

fn main() -> UnlockResult<()> {
    let cli = Cli::parse();

    match &cli.command {
        Commands::Read(args) => {
            let file_and_path = std::path::Path::new(&args.filename);
            let extension = file_and_path
                .extension()
                .and_then(|s| s.to_str())
                .map(str::to_lowercase);

            match extension.as_deref() {
                Some("xls") => {
                    let mut file = cfb::open(file_and_path).map_err(UnlockError::CFBOpen)?;
                    let project = file.open_stream(consts::CFB_VBA_PATH)?;
                    print_info(project, args.decode)?;
                }
                Some("xlsm" | "xlsb") => {
                    let zipfile = std::fs::File::open(file_and_path)?;
                    let mut archive = zip::ZipArchive::new(zipfile)?;
                    let Ok(mut file) = archive.by_name(consts::ZIP_VBA_PATH) else {
                        return Err(UnlockError::NoVBAFile);
                    };

                    // Read the uncompressed bytes of the vbaProject.bin file into an in-memory cursor
                    // Need this as ZipFile does not implement Seek, so we cannot call open_stream
                    // on a CompoundFile that is built directly off the ZipFile
                    let mut buffer = Vec::with_capacity(1024);
                    let _ = file.read_to_end(&mut buffer);
                    let raw = Cursor::new(buffer);

                    let mut vba = CompoundFile::open(raw).map_err(UnlockError::CFBOpen)?;
                    let project = vba.open_stream(consts::PROJECT_PATH)?;
                    print_info(project, args.decode)?;
                }
                Some("xlsx") => {
                    return Err(UnlockError::XlsX(
                        file_and_path.to_string_lossy().to_string(),
                    ))
                }
                _ => {
                    return Err(UnlockError::NotExcel(
                        file_and_path.to_string_lossy().to_string(),
                    ))
                }
            }
        }
        Commands::Remove(_inplace) => {
            println!("Not yet built. Sorry");
        }
    }

    Ok(())
}
