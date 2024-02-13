#![warn(clippy::all, clippy::pedantic, clippy::nursery)]

mod consts;
mod error;
mod read;

use cfb::CompoundFile;
use clap::{Args, Parser, Subcommand};
use std::io::{prelude::*, Cursor};
use std::path::Path;

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

enum XlType {
    Old,
    New,
}

fn main() -> UnlockResult<()> {
    let cli = Cli::parse();
    let (filename, version) = get_file(&cli)?;
    match (&cli.command, version) {
        (Commands::Read(args), XlType::Old) => {
            let mut file = cfb::open(filename).map_err(UnlockError::CFBOpen)?;
            let project = file.open_stream(consts::CFB_VBA_PATH)?;
            print_info(project, args.decode)?;
        }
        (Commands::Read(args), XlType::New) => {
            let zipfile = std::fs::File::open(filename)?;
            let mut archive = zip::ZipArchive::new(zipfile)?;
            let Ok(mut vba_file) = archive.by_name(consts::ZIP_VBA_PATH) else {
                return Err(UnlockError::NoVBAFile);
            };

            // Read the uncompressed bytes of the vbaProject.bin file into an in-memory cursor
            // Need this as ZipFile does not implement Seek, so we cannot call open_stream
            // on a CompoundFile that is built directly off the ZipFile
            let mut buffer = Vec::with_capacity(1024);
            let _ = vba_file.read_to_end(&mut buffer);
            let vba_raw = Cursor::new(buffer);

            let mut vba = CompoundFile::open(vba_raw).map_err(UnlockError::CFBOpen)?;
            let project = vba.open_stream(consts::PROJECT_PATH)?;
            print_info(project, args.decode)?;
        }
        (Commands::Remove(_), _) => {
            println!("Not yet built. Sorry");
        }
    }

    Ok(())
}

fn get_file(cli: &Cli) -> UnlockResult<(&Path, XlType)> {
    let filename = match &cli.command {
        Commands::Read(a) => a.filename.as_str(),
        Commands::Remove(a) => a.filename.as_str(),
    };
    let filename = std::path::Path::new(filename);
    let extension = filename
        .extension()
        .and_then(|s| s.to_str())
        .map(str::to_lowercase);

    let version = match extension.as_deref() {
        Some("xls") => XlType::Old,
        Some("xlsm" | "xlsb") => XlType::New,
        Some("xlsx") => return Err(UnlockError::XlsX(filename.to_string_lossy().to_string())),
        _ => {
            return Err(UnlockError::NotExcel(
                filename.to_string_lossy().to_string(),
            ))
        }
    };

    Ok((filename, version))
}
