#![warn(clippy::all, clippy::pedantic, clippy::nursery)]

mod consts;
mod error;
mod read;

use clap::{Args, Parser, Subcommand};
use std::io::{prelude::*, Cursor};

use error::UnlockError;
use error::UnlockResult;
use read::print_info;

use cfb::{CompoundFile, Stream};
pub type InMemCFB = CompoundFile<Cursor<Vec<u8>>>;
pub type InMemStream = Stream<Cursor<Vec<u8>>>;

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

fn main() {
    let cli = Cli::parse();

    match &cli.command {
        Commands::Read(args) => {
            let mut vba = match get_vba(&args.filename) {
                Ok(p) => p,
                Err(e) => {
                    eprintln!("{e}");
                    std::process::exit(1);
                }
            };
            let project = vba.open_stream(consts::PROJECT_PATH).unwrap();
            if let Err(e) = print_info(project) {
                eprintln!("{e}");
                std::process::exit(1);
            }
        }
        Commands::Remove(_inplace) => {
            println!("Not yet built. Sorry");
        }
    }

    std::process::exit(0);
}

fn get_vba(path: &str) -> UnlockResult<InMemCFB> {
    let fname = std::path::Path::new(path);
    let zipfile = std::fs::File::open(fname)?;

    let mut archive = zip::ZipArchive::new(zipfile)?;

    let Ok(mut file) = archive.by_name(consts::VBA_PATH) else {
        return Err(UnlockError::NoVBAFile);
    };

    // Read the uncompressed bytes of the vbaProject.bin file into an in-memory cursor
    let mut buffer = Vec::with_capacity(1024);
    let _ = file.read_to_end(&mut buffer);
    let raw = std::io::Cursor::new(buffer);

    CompoundFile::open(raw).map_err(UnlockError::CFBOpen)
}
