use crate::consts;
use crate::error::UnlockError;
use crate::error::UnlockResult;
use crate::read::zip_to_raw_vba;
use cfb::Stream;
use std::fs::File;
use std::io::{BufRead, Write};
use std::path::Path;
use std::path::PathBuf;

/// Remove the VBA protection from an Excel file
/// This is the version for Excel files since 2003 i.e. xlsm and xlsb
///
/// The inplace flag, if set to true, will overwrite the source file with a modified unlocked
/// version. It is recommended to take a back-up of the file before doing this as the tool is
/// relatively new and untested. It may corrupt your file.
///
/// Alternatively, pass false for the inplace flag to get a copy of the source file. It will have
/// the same name as the source file, but have '_unlocked' appended to the filename.
///
/// # Errors
/// Will return an error in the following situations:
/// - The file cannot be opened
/// - The file is cannot be opened as a zip file: Excel files since 2003 are really zip files. The
/// contents within the zip file changes depending on the Excel file format used: xlsx, xlsm, xlsb
/// - There is no VBA file within the zip archive, found at "/xl/vbaProject.bin". Note that an
/// xlsm file saved with no macros will be missing this file, as will any xlsx file. In the former
/// case, the code really ought to handle the "error" more gracefully
/// - The VBA file within the archive cannot be opened as a [Compound File Binary (CFB)](https://learn.microsoft.com/en-us/openspecs/windows_protocols/MS-CFB/53989ce4-7b05-4f8d-829b-d08d6148375b).
/// This file format stores the data of a file as a mini file system. The data of each "file"
/// within the overall file is stored as streams. These streams are written to 512 byte sectors, or
/// 64 byte chunks of the mini-stream. In either case, the sectors or the mini-stream, the stream
/// is not guaranteed to be written to contiguous memory, so it is important that the file is
/// properly opened as a CFB file in order to read the streams correctly
/// - The [PROJECT stream](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/cc848a02-6f87-49a4-ad93-6edb3103f593),
/// which holds the VBA locked status, cannot be found within the overall VBA CFB file
/// - The updated project stream cannot be written back to the CFB file
/// - An updated zip file cannot be created
/// - The updated VBA CFB file cannot be written to the new zip file
/// - The rest of the source zip file cannot be copied across as raw to the new zip file
/// - If being run inplace, the new zip file cannot be copied back over the original
pub fn xl(filename: &Path, inplace: bool) -> UnlockResult<()> {
    let zipfile = File::open(filename)?;
    let mut archive = zip::ZipArchive::new(zipfile)?;
    let vba_raw = zip_to_raw_vba(&mut archive)?;

    // Replace the VBA CFB file with an unlocked project
    // Strip back out to a Vec of bytes as this is what's needed to write to the zip file
    let mut vba = cfb::CompoundFile::open(vba_raw).map_err(UnlockError::CFBOpen)?;
    let project = vba.open_stream(consts::PROJECT_PATH)?;
    let replacement = unlocked_project(project)?;
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
                new_archive.start_file(consts::ZIP_VBA_PATH, zip::write::FileOptions::default())?;
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
    if inplace {
        std::fs::rename(new_filename, filename)?;
    }

    Ok(())
}

/// Remove the VBA protection from an Excel file
/// This is the version for Excel files between 1997 & 2003 i.e. xls
///
/// The inplace flag, if set to true, will overwrite the source file with a modified unlocked
/// version. It is recommended to take a back-up of the file before doing this as the tool is
/// relatively new and untested. It may corrupt your file.
///
/// Alternatively, pass false for the inplace flag to get a copy of the source file. It will have
/// the same name as the source file, but have '_unlocked' appended to the filename.
///
/// # Errors
/// Will return an error in the following situations:
/// - The file cannot be copied (for not inplace only) or opened for read/write
/// - The file cannot be opened as a [Compound File Binary (CFB)](https://learn.microsoft.com/en-us/openspecs/windows_protocols/MS-CFB/53989ce4-7b05-4f8d-829b-d08d6148375b).
/// This file format stores the data of a file as a mini file system. The data of each "file"
/// within the overall file is stored as streams. These streams are written to 512 byte sectors, or
/// 64 byte chunks of the mini-stream. In either case, the sectors or the mini-stream, the stream
/// is not guaranteed to be written to contiguous memory, so it is important that the file is
/// properly opened as a CFB file in order to read the streams correctly
/// - The [PROJECT stream](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/cc848a02-6f87-49a4-ad93-6edb3103f593),
/// which holds the VBA locked status, cannot be found within the overall VBA CFB file
/// - The updated project stream cannot be written back to the CFB file
pub fn xl_97(filename: &Path, inplace: bool) -> UnlockResult<()> {
    let mut file = if inplace {
        cfb::open_rw(filename).map_err(UnlockError::CFBOpen)?
    } else {
        let new_file = replacement_filename(filename)?;
        std::fs::copy(filename, &new_file)?;
        cfb::open_rw(new_file).map_err(UnlockError::CFBOpen)?
    };
    let project = file.open_stream(consts::CFB_VBA_PATH)?;
    let replacement = unlocked_project(project)?;
    let mut project = file.create_stream(consts::CFB_VBA_PATH)?;
    Ok(project.write_all(&replacement)?)
}

fn unlocked_project<T: std::io::Read + std::io::Seek>(
    mut project: Stream<T>,
) -> UnlockResult<Vec<u8>> {
    let mut line = Vec::new();
    let mut output = Vec::new();

    while project.read_until(b'\n', &mut line)? > 0 {
        match line.get(0..5) {
            Some(&[b'I', b'D', b'=', b'"', b'{']) => {
                output.extend_from_slice(consts::UNLOCKED_ID.as_bytes());
            }
            Some(&[b'C', b'M', b'G', b'=', b'"']) => {
                output.extend_from_slice(consts::UNLOCKED_CMG.as_bytes());
            }
            Some(&[b'D', b'P', b'B', b'=', b'"']) => {
                output.extend_from_slice(consts::UNLOCKED_DPB.as_bytes());
            }
            Some(&[b'G', b'C', b'=', b'"', _]) => {
                output.extend_from_slice(consts::UNLOCKED_GC.as_bytes());
            }
            _ => output.extend_from_slice(&line),
        }
        line.clear();
    }

    Ok(output)
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
