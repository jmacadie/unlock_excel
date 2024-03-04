use std::fs::File;
use std::io::{Cursor, Read};
use std::path::Path;

use crate::consts;
use crate::error::{UnlockError, UnlockResult};
use crate::ovba::records::project::{Password, Project};
use sha1::{Digest, Sha1};
use zip::ZipArchive;

/// Print the VBA project locked status to standard out.
/// This is the version for Excel files since 2003 i.e. xlsm and xlsb
///
/// The decode flag, if set to true, will trigger an attempt to decode a SHA hashed password. This
/// is done by testing against [a list of 1.7 million common passwords](https://github.com/openwall/john/blob/bleeding-jumbo/run/password.lst)
///
/// # Errors
/// Will return an error in the following situations:
/// - The file cannot be opened
/// - The file is cannot be opened as a zip file: Excel files since 2003 are really zip files. The
/// contents within the zip file changes depending on the Excel file format used: xlsx, xlsm, xlsb
/// - If there is no VBA file within the zip archive, found at "/xl/vbaProject.bin". Note that an
/// xlsm file saved with no macros will be missing this file, as will any xlsx file. In the former
/// case, the code really ought to handle the "error" more gracefully
/// - If the VBA file within the archive cannot be opened as a [Compound File Binary (CFB)](https://learn.microsoft.com/en-us/openspecs/windows_protocols/MS-CFB/53989ce4-7b05-4f8d-829b-d08d6148375b).
/// This file format stores the data of a file as a mini file system. The data of each "file"
/// within the overall file is stored as streams. These streams are written to 512 byte sectors, or
/// 64 byte chunks of the mini-stream. In either case, the sectors or the mini-stream, the stream
/// is not guaranteed to be written to contiguous memory, so it is important that the file is
/// properly opened as a CFB file in order to read the streams correctly
/// - If the [PROJECT stream](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/cc848a02-6f87-49a4-ad93-6edb3103f593),
/// which holds the VBA locked status, cannot be found within the overall VBA CFB file
/// - If the [PROJECT stream cannot be parsed](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/cc848a02-6f87-49a4-ad93-6edb3103f593)
/// into its constituent parts correctly
pub fn print_xl(filename: &Path, decode: bool) -> UnlockResult<()> {
    let (project, decoded_password) = xl_project(filename, decode)?;
    print_info(&project, decode, decoded_password);
    Ok(())
}

/// Parse an Excel file into an [`ovba::records::project::Project`].
/// This is exposed to allow for integration testing.
/// This is the version for Excel files since 2003 i.e. xlsm and xlsb
///
/// The decode flag, if set to true, will trigger an attempt to decode a SHA hashed password. This
/// is done by testing against [a list of 1.7 million common passwords](https://github.com/openwall/john/blob/bleeding-jumbo/run/password.lst)
///
/// # Errors
/// Will return an error in the following situations:
/// - The file cannot be opened
/// - The file is cannot be opened as a zip file: Excel files since 2003 are really zip files. The
/// contents within the zip file changes depending on the Excel file format used: xlsx, xlsm, xlsb
/// - If there is no VBA file within the zip archive, found at "/xl/vbaProject.bin". Note that an
/// xlsm file saved with no macros will be missing this file, as will any xlsx file. In the former
/// case, the code really ought to handle the "error" more gracefully
/// - If the VBA file within the archive cannot be opened as a [Compound File Binary (CFB)](https://learn.microsoft.com/en-us/openspecs/windows_protocols/MS-CFB/53989ce4-7b05-4f8d-829b-d08d6148375b).
/// This file format stores the data of a file as a mini file system. The data of each "file"
/// within the overall file is stored as streams. These streams are written to 512 byte sectors, or
/// 64 byte chunks of the mini-stream. In either case, the sectors or the mini-stream, the stream
/// is not guaranteed to be written to contiguous memory, so it is important that the file is
/// properly opened as a CFB file in order to read the streams correctly
/// - If the [PROJECT stream](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/cc848a02-6f87-49a4-ad93-6edb3103f593),
/// which holds the VBA locked status, cannot be found within the overall VBA CFB file
/// - If the [PROJECT stream cannot be parsed](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/cc848a02-6f87-49a4-ad93-6edb3103f593)
/// into its constituent parts correctly
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

/// Print the VBA project locked status to standard out.
/// This is the version for Excel files between 1997 & 2003 i.e. xls
///
/// The decode flag, if set to true, will trigger an attempt to decode a SHA hashed password. This
/// is done by testing against [a list of 1.7 million common passwords](https://github.com/openwall/john/blob/bleeding-jumbo/run/password.lst)
///
/// # Errors
/// Will return an error in the following situations:
/// - The file cannot be opened
/// - If the file cannot be opened as a [Compound File Binary (CFB)](https://learn.microsoft.com/en-us/openspecs/windows_protocols/MS-CFB/53989ce4-7b05-4f8d-829b-d08d6148375b).
/// This file format stores the data of a file as a mini file system. The data of each "file"
/// within the overall file is stored as streams. These streams are written to 512 byte sectors, or
/// 64 byte chunks of the mini-stream. In either case, the sectors or the mini-stream, the stream
/// is not guaranteed to be written to contiguous memory, so it is important that the file is
/// properly opened as a CFB file in order to read the streams correctly
/// - If the [PROJECT stream](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/cc848a02-6f87-49a4-ad93-6edb3103f593),
/// which holds the VBA locked status, cannot be found within the overall CFB file
/// - If the [PROJECT stream cannot be parsed](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/cc848a02-6f87-49a4-ad93-6edb3103f593)
/// into its constituent parts correctly
pub fn print_xl_97(filename: &Path, decode: bool) -> UnlockResult<()> {
    let (project, decoded_password) = xl_97_project(filename, decode)?;
    print_info(&project, decode, decoded_password);
    Ok(())
}

/// Parse an Excel file into an [`ovba::records::project::Project`].
/// This is exposed to allow for integration testing.
/// This is the version for Excel files between 1997 & 2003 i.e. xls
///
/// The decode flag, if set to true, will trigger an attempt to decode a SHA hashed password. This
/// is done by testing against [a list of 1.7 million common passwords](https://github.com/openwall/john/blob/bleeding-jumbo/run/password.lst)
///
/// # Errors
/// Will return an error in the following situations:
/// - The file cannot be opened
/// - If the file cannot be opened as a [Compound File Binary (CFB)](https://learn.microsoft.com/en-us/openspecs/windows_protocols/MS-CFB/53989ce4-7b05-4f8d-829b-d08d6148375b).
/// This file format stores the data of a file as a mini file system. The data of each "file"
/// within the overall file is stored as streams. These streams are written to 512 byte sectors, or
/// 64 byte chunks of the mini-stream. In either case, the sectors or the mini-stream, the stream
/// is not guaranteed to be written to contiguous memory, so it is important that the file is
/// properly opened as a CFB file in order to read the streams correctly
/// - If the [PROJECT stream](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/cc848a02-6f87-49a4-ad93-6edb3103f593),
/// which holds the VBA locked status, cannot be found within the overall CFB file
/// - If the [PROJECT stream cannot be parsed](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/cc848a02-6f87-49a4-ad93-6edb3103f593)
/// into its constituent parts correctly
pub fn xl_97_project(filename: &Path, decode: bool) -> UnlockResult<(Project, Option<String>)> {
    let mut file = cfb::open(filename).map_err(UnlockError::CFBOpen)?;
    let project_stream = file.open_stream(consts::CFB_VBA_PATH)?;
    let project = Project::from_stream(project_stream)?;
    let decoded_password = decode
        .then(|| try_solve_password(project.password()))
        .flatten();
    Ok((project, decoded_password))
}

/// Read the uncompressed bytes of the vbaProject.bin file into an in-memory cursor
///
/// Need this as `ZipFile` does not implement Seek, so we cannot call `open_stream`
/// on a `CompoundFile` that is built directly off the `ZipFile`
pub(crate) fn zip_to_raw_vba<R: std::io::Read + std::io::Seek>(
    zip: &mut ZipArchive<R>,
) -> UnlockResult<Cursor<Vec<u8>>> {
    let Ok(mut vba_file) = zip.by_name(consts::ZIP_VBA_PATH) else {
        return Err(UnlockError::NoVBAFile);
    };

    let mut buffer = Vec::with_capacity(1024);
    let _ = vba_file.read_to_end(&mut buffer);
    Ok(Cursor::new(buffer))
}

/// Internal function to print the results of the Project stuct to stdout consistently
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
