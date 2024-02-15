#![warn(clippy::all, clippy::pedantic, clippy::nursery)]

mod consts;
mod error;
mod read;
mod remove;

use clap::{Args, Parser, Subcommand};
use std::fs::File;
use std::io::{prelude::*, Cursor};
use std::path::{Path, PathBuf};

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
            let zipfile = File::open(filename)?;
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

            let mut vba = cfb::CompoundFile::open(vba_raw).map_err(UnlockError::CFBOpen)?;
            let project = vba.open_stream(consts::PROJECT_PATH)?;
            print_info(project, args.decode)?;
        }
        (Commands::Remove(args), XlType::Old) => {
            let mut file = if args.inplace {
                cfb::open_rw(filename).map_err(UnlockError::CFBOpen)?
            } else {
                let new_file = replacement_filename(filename)?;
                std::fs::copy(filename, &new_file)?;
                cfb::open_rw(new_file).map_err(UnlockError::CFBOpen)?
            };
            let project = file.open_stream(consts::CFB_VBA_PATH)?;
            let replacement = remove::unlocked_project(project)?;
            let mut project = file.create_stream(consts::CFB_VBA_PATH)?;
            project.write_all(&replacement)?;
        }
        (Commands::Remove(args), XlType::New) => {
            let zipfile = File::open(filename)?;
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
            drop(vba_file);

            // Replace the VBA CFB file with an unlocked project
            // Strip back out to a Vec of bytes as this is what's needed to write to the zip file
            let mut vba = cfb::CompoundFile::open(vba_raw).map_err(UnlockError::CFBOpen)?;
            let project = vba.open_stream(consts::PROJECT_PATH)?;
            let replacement = remove::unlocked_project(project)?;
            let mut project = vba.create_stream(consts::PROJECT_PATH)?;
            project.write_all(&replacement)?;
            project.flush()?;
            let vba_inner = vba.into_inner().into_inner();

            // Open a new, empty archive for writing to
            let new_filename = replacement_filename(filename)?;
            let new_file = File::create(&new_filename)?;
            let mut new_archive = zip::ZipWriter::new(new_file);

            // Loop through the original archive:
            //  - Write the VBA file from our updated vec of bytes
            //  - Copy everything else across as raw, which saves the bother of decoding it
            // The end effect is to have a new archive, which is a clone of the original,
            // save for the VBA file which has been rewritten
            let target: &Path = consts::ZIP_VBA_PATH.as_ref();
            for i in 0..archive.len() {
                let file = archive.by_index_raw(i)?;
                match file.enclosed_name() {
                    Some(p) if p == target => {
                        new_archive
                            .start_file(consts::ZIP_VBA_PATH, zip::write::FileOptions::default())?;
                        new_archive.write_all(&vba_inner)?;
                        new_archive.flush()?;
                    }
                    _ => new_archive.raw_copy_file(file)?,
                }
            }
            new_archive.finish()?;

            drop(archive);
            drop(new_archive);

            // If we're doing this in place then overwrite the original with the new
            if args.inplace {
                std::fs::rename(new_filename, filename)?;
            }
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

    match extension.as_deref() {
        Some("xls") => Ok((filename, XlType::Old)),
        Some("xlsm" | "xlsb") => Ok((filename, XlType::New)),
        Some("xlsx") => Err(UnlockError::XlsX(filename.to_string_lossy().to_string())),
        _ => Err(UnlockError::NotExcel(
            filename.to_string_lossy().to_string(),
        )),
    }
}

fn replacement_filename(source: &Path) -> UnlockResult<PathBuf> {
    let mut new = PathBuf::from(source);
    let mut stem = source
        .file_stem()
        .ok_or(UnlockError::NotExcel(source.to_string_lossy().to_string()))?
        .to_owned();
    stem.push("_unlocked");
    new.set_file_name(stem);
    let ext = source
        .extension()
        .ok_or(UnlockError::NotExcel(source.to_string_lossy().to_string()))?;
    new.set_extension(ext);
    Ok(new)
}
