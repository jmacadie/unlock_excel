#![warn(clippy::all, clippy::pedantic, clippy::nursery)]

mod consts;
mod error;
mod ovba;
mod read;
mod remove;

use clap::{Args, Parser, Subcommand};
use std::path::Path;

use crate::error::UnlockError;
use crate::error::UnlockResult;

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
        (Commands::Read(args), XlType::Old) => read::xl_97(filename, args.decode)?,
        (Commands::Read(args), XlType::New) => read::xl(filename, args.decode)?,
        (Commands::Remove(args), XlType::Old) => remove::xl_97(filename, args.inplace)?,
        (Commands::Remove(args), XlType::New) => remove::xl(filename, args.inplace)?,
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

    match extension.as_deref() {
        Some("xls") => Ok((filename, XlType::Old)),
        Some("xlsm" | "xlsb") => Ok((filename, XlType::New)),
        Some("xlsx") => Err(UnlockError::XlsX(filename.to_string_lossy().to_string())),
        _ => Err(UnlockError::NotExcel(
            filename.to_string_lossy().to_string(),
        )),
    }
}
