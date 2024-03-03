use std::fs::File;
use std::io::{Cursor, Read};
use std::path::Path;

use crate::consts;
use crate::error::{UnlockError, UnlockResult};
use crate::ovba::records::project::{Password, Project};
use sha1::{Digest, Sha1};
use zip::ZipArchive;

pub fn print_xl(filename: &Path, decode: bool) -> UnlockResult<()> {
    let (project, decoded_password) = xl_project(filename, decode)?;
    print_info(&project, decode, decoded_password);
    Ok(())
}

pub fn xl_project(filename: &Path, decode: bool) -> UnlockResult<(Project, Option<String>)> {
    let zipfile = File::open(filename)?;
    let mut archive = zip::ZipArchive::new(zipfile)?;
    let vba_raw = zip_to_raw_vba(&mut archive)?;
    let mut vba_cfb = cfb::CompoundFile::open(vba_raw).map_err(UnlockError::CFBOpen)?;
    let project_stream = vba_cfb.open_stream(consts::PROJECT_PATH)?;
    let project = Project::from_stream(project_stream)?;
    let decoded_password = decode
        .then(|| try_solve_password(project.password()))
        .flatten();
    Ok((project, decoded_password))
}

pub fn print_xl_97(filename: &Path, decode: bool) -> UnlockResult<()> {
    let (project, decoded_password) = xl_97_project(filename, decode)?;
    print_info(&project, decode, decoded_password);
    Ok(())
}

pub fn xl_97_project(filename: &Path, decode: bool) -> UnlockResult<(Project, Option<String>)> {
    let mut file = cfb::open(filename).map_err(UnlockError::CFBOpen)?;
    let project_stream = file.open_stream(consts::CFB_VBA_PATH)?;
    let project = Project::from_stream(project_stream)?;
    let decoded_password = decode
        .then(|| try_solve_password(project.password()))
        .flatten();
    Ok((project, decoded_password))
}

pub fn zip_to_raw_vba<R: std::io::Read + std::io::Seek>(
    zip: &mut ZipArchive<R>,
) -> UnlockResult<Cursor<Vec<u8>>> {
    let Ok(mut vba_file) = zip.by_name(consts::ZIP_VBA_PATH) else {
        return Err(UnlockError::NoVBAFile);
    };

    // Read the uncompressed bytes of the vbaProject.bin file into an in-memory cursor
    // Need this as ZipFile does not implement Seek, so we cannot call open_stream
    // on a CompoundFile that is built directly off the ZipFile
    let mut buffer = Vec::with_capacity(1024);
    let _ = vba_file.read_to_end(&mut buffer);
    Ok(Cursor::new(buffer))
}

fn print_info(p: &Project, decode: bool, decoded: Option<String>) {
    if p.is_locked() {
        match p.password() {
            Password::None => {
                println!("😕 The VBA is locked with no password");
                println!("This should never happen 🤷");
            }
            Password::Hash(salt, hash) => {
                println!("🔐 The VBA is locked");
                println!();
                println!("The password (+ a salt) has been stored as a SHA1 hash:");
                print!("Hash: ");
                for byte in hash {
                    print!("{byte:02x}");
                }
                println!();
                print!("Salt: ");
                for byte in salt {
                    print!("{byte:02x}");
                }
                println!();
                match (decode, decoded) {
                    (true, Some(s)) => {
                        println!();
                        println!("✅ Was able to decode this weak password: {s}");
                    }
                    (true, None) => {
                        println!();
                        println!("❌ Was unable to decode this password");
                        println!("You can just remove the password with `unlock_excel remove FILENAME`, which will always work");
                    }
                    (false, _) => (),
                }
            }
            Password::Plain(text) => {
                println!("🔒 The VBA is locked");
                println!();
                println!("The password has been stored as plain-text though: {text}");
            }
        }
    } else {
        println!("🔓 The VBA is not locked");
        println!("You can freely open it 🥳");
    }
}

fn try_solve_password(p: &Password) -> Option<String> {
    match p {
        Password::Hash(salt, hash) => {
            let words = include_str!("password.lst");
            let mut hasher = Sha1::new();
            for trial in words.lines() {
                let mut salted: Vec<u8> = trial.as_bytes().to_owned();
                salted.extend_from_slice(salt);
                hasher.update(salted);
                if hasher.finalize_reset()[..] == *hash {
                    return Some(trial.to_owned());
                }
            }
            None
        }
        _ => None,
    }
}
