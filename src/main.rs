#![warn(clippy::all, clippy::pedantic, clippy::nursery)]

mod consts;
mod error;
mod read;

use cfb::{CompoundFile, Stream};
use consts::WORDS;
use error::UnlockResult;
use std::io::{prelude::*, Cursor};

use crate::error::UnlockError;
use crate::read::print_info;

pub type InMemCFB = CompoundFile<Cursor<Vec<u8>>>;
pub type InMemStream = Stream<Cursor<Vec<u8>>>;

fn main() {
    std::process::exit(real_main());
}

fn real_main() -> i32 {
    const PROJECT_PATH: &str = "/PROJECT";

    let args: Vec<_> = std::env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: {} <filename>", args[0]);
        return 1;
    }

    let mut cfb_file = match get_vba(&args[1]) {
        Ok(f) => f,
        Err(e) => {
            eprintln!("{e}");
            return 1;
        }
    };

    let project = match cfb_file.open_stream(PROJECT_PATH) {
        Ok(s) => s,
        Err(e) => {
            eprintln!("Cannot open the Project Stream: {e}");
            return 1;
        }
    };

    if let Err(e) = print_info(project) {
        eprintln!("{e}");
        return 1;
    }

    0
}

fn get_vba(path: &str) -> UnlockResult<InMemCFB> {
    let fname = std::path::Path::new(path);
    let zipfile = std::fs::File::open(fname)?;

    let mut archive = zip::ZipArchive::new(zipfile)?;

    let Ok(mut file) = archive.by_name("xl/vbaProject.bin") else {
        return Err(UnlockError::NoVBAFile);
    };

    // Read the uncompressed bytes of the vbaProject.bin file into an in-memory cursor
    let mut buffer = Vec::with_capacity(1024);
    let _ = file.read_to_end(&mut buffer);
    let raw = std::io::Cursor::new(buffer);

    CompoundFile::open(raw).map_err(UnlockError::CFBOpen)
}
