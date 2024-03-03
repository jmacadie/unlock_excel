use std::path::{Path, PathBuf};
use unlock_excel::read;
use unlock_excel::remove::{xl, xl_97};

/*
* XLSM
* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

#[test]
fn remove_unlocked_copy_xlsm() {
    let file = "tests/data/xlsm/Unlocked_with_macro.xlsm";
    let (temp_dir, temp_file) = create_temp_dir(&file, 1);
    let replacement = replacement_filename(&temp_file);
    xl(Path::new(&temp_file), false).unwrap();
    let (p, _) = read::xl_project(&replacement, false).unwrap();
    assert!(!p.is_locked());
    let _ = std::fs::remove_dir_all(temp_dir);
}

#[test]
fn remove_locked_1_copy_xlsm() {
    let file = "tests/data/xlsm/Locked_with_macro.xlsm";
    let (temp_dir, temp_file) = create_temp_dir(&file, 2);
    let replacement = replacement_filename(&temp_file);
    xl(Path::new(&temp_file), false).unwrap();
    let (p, _) = read::xl_project(&replacement, false).unwrap();
    assert!(!p.is_locked());
    let _ = std::fs::remove_dir_all(temp_dir);
}

#[test]
fn remove_locked_2_copy_xlsm() {
    let file = "tests/data/xlsm/Locked_with_macro_and_complex_password.xlsm";
    let (temp_dir, temp_file) = create_temp_dir(&file, 3);
    let replacement = replacement_filename(&temp_file);
    xl(Path::new(&temp_file), false).unwrap();
    let (p, _) = read::xl_project(&replacement, false).unwrap();
    assert!(!p.is_locked());
    let _ = std::fs::remove_dir_all(temp_dir);
}

#[test]
fn remove_unlocked_inplace_xlsm() {
    let file = "tests/data/xlsm/Unlocked_with_macro.xlsm";
    let (temp_dir, temp_file) = create_temp_dir(&file, 4);
    xl(Path::new(&temp_file), true).unwrap();
    let (p, _) = read::xl_project(&temp_file, false).unwrap();
    assert!(!p.is_locked());
    let _ = std::fs::remove_dir_all(temp_dir);
}

#[test]
fn remove_locked_1_inplace_xlsm() {
    let file = "tests/data/xlsm/Locked_with_macro.xlsm";
    let (temp_dir, temp_file) = create_temp_dir(&file, 5);
    xl(Path::new(&temp_file), true).unwrap();
    let (p, _) = read::xl_project(&temp_file, false).unwrap();
    assert!(!p.is_locked());
    let _ = std::fs::remove_dir_all(temp_dir);
}

#[test]
fn remove_locked_2_inplace_xlsm() {
    let file = "tests/data/xlsm/Locked_with_macro_and_complex_password.xlsm";
    let (temp_dir, temp_file) = create_temp_dir(&file, 6);
    xl(Path::new(&temp_file), true).unwrap();
    let (p, _) = read::xl_project(&temp_file, false).unwrap();
    assert!(!p.is_locked());
    let _ = std::fs::remove_dir_all(temp_dir);
}

/*
* XLSB
* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

#[test]
fn remove_unlocked_copy_xlsb() {
    let file = "tests/data/xlsb/Unlocked_with_macro.xlsb";
    let (temp_dir, temp_file) = create_temp_dir(&file, 1);
    let replacement = replacement_filename(&temp_file);
    xl(Path::new(&temp_file), false).unwrap();
    let (p, _) = read::xl_project(&replacement, false).unwrap();
    assert!(!p.is_locked());
    let _ = std::fs::remove_dir_all(temp_dir);
}

#[test]
fn remove_locked_1_copy_xlsb() {
    let file = "tests/data/xlsb/Locked_with_macro.xlsb";
    let (temp_dir, temp_file) = create_temp_dir(&file, 2);
    let replacement = replacement_filename(&temp_file);
    xl(Path::new(&temp_file), false).unwrap();
    let (p, _) = read::xl_project(&replacement, false).unwrap();
    assert!(!p.is_locked());
    let _ = std::fs::remove_dir_all(temp_dir);
}

#[test]
fn remove_locked_2_copy_xlsb() {
    let file = "tests/data/xlsb/Locked_with_macro_and_complex_password.xlsb";
    let (temp_dir, temp_file) = create_temp_dir(&file, 3);
    let replacement = replacement_filename(&temp_file);
    xl(Path::new(&temp_file), false).unwrap();
    let (p, _) = read::xl_project(&replacement, false).unwrap();
    assert!(!p.is_locked());
    let _ = std::fs::remove_dir_all(temp_dir);
}

#[test]
fn remove_unlocked_inplace_xlsb() {
    let file = "tests/data/xlsb/Unlocked_with_macro.xlsb";
    let (temp_dir, temp_file) = create_temp_dir(&file, 4);
    xl(Path::new(&temp_file), true).unwrap();
    let (p, _) = read::xl_project(&temp_file, false).unwrap();
    assert!(!p.is_locked());
    let _ = std::fs::remove_dir_all(temp_dir);
}

#[test]
fn remove_locked_1_inplace_xlsb() {
    let file = "tests/data/xlsb/Locked_with_macro.xlsb";
    let (temp_dir, temp_file) = create_temp_dir(&file, 5);
    xl(Path::new(&temp_file), true).unwrap();
    let (p, _) = read::xl_project(&temp_file, false).unwrap();
    assert!(!p.is_locked());
    let _ = std::fs::remove_dir_all(temp_dir);
}

#[test]
fn remove_locked_2_inplace_xlsb() {
    let file = "tests/data/xlsb/Locked_with_macro_and_complex_password.xlsb";
    let (temp_dir, temp_file) = create_temp_dir(&file, 6);
    xl(Path::new(&temp_file), true).unwrap();
    let (p, _) = read::xl_project(&temp_file, false).unwrap();
    assert!(!p.is_locked());
    let _ = std::fs::remove_dir_all(temp_dir);
}

/*
* XLS
* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

#[test]
fn remove_unlocked_copy_xls() {
    let file = "tests/data/xls/Unlocked_with_macro.xls";
    let (temp_dir, temp_file) = create_temp_dir(&file, 1);
    let replacement = replacement_filename(&temp_file);
    xl_97(Path::new(&temp_file), false).unwrap();
    let (p, _) = read::xl_97_project(&replacement, false).unwrap();
    assert!(!p.is_locked());
    let _ = std::fs::remove_dir_all(temp_dir);
}

#[test]
fn remove_locked_1_copy_xls() {
    let file = "tests/data/xls/Locked_with_macro.xls";
    let (temp_dir, temp_file) = create_temp_dir(&file, 2);
    let replacement = replacement_filename(&temp_file);
    xl_97(Path::new(&temp_file), false).unwrap();
    let (p, _) = read::xl_97_project(&replacement, false).unwrap();
    assert!(!p.is_locked());
    let _ = std::fs::remove_dir_all(temp_dir);
}

#[test]
fn remove_locked_2_copy_xls() {
    let file = "tests/data/xls/Locked_with_macro_and_complex_password.xls";
    let (temp_dir, temp_file) = create_temp_dir(&file, 3);
    let replacement = replacement_filename(&temp_file);
    xl_97(Path::new(&temp_file), false).unwrap();
    let (p, _) = read::xl_97_project(&replacement, false).unwrap();
    assert!(!p.is_locked());
    let _ = std::fs::remove_dir_all(temp_dir);
}

#[test]
fn remove_unlocked_inplace_xls() {
    let file = "tests/data/xls/Unlocked_with_macro.xls";
    let (temp_dir, temp_file) = create_temp_dir(&file, 4);
    xl_97(Path::new(&temp_file), true).unwrap();
    let (p, _) = read::xl_97_project(&temp_file, false).unwrap();
    assert!(!p.is_locked());
    let _ = std::fs::remove_dir_all(temp_dir);
}

#[test]
fn remove_locked_1_inplace_xls() {
    let file = "tests/data/xls/Locked_with_macro.xls";
    let (temp_dir, temp_file) = create_temp_dir(&file, 5);
    xl_97(Path::new(&temp_file), true).unwrap();
    let (p, _) = read::xl_97_project(&temp_file, false).unwrap();
    assert!(!p.is_locked());
    let _ = std::fs::remove_dir_all(temp_dir);
}

#[test]
fn remove_locked_2_inplace_xls() {
    let file = "tests/data/xls/Locked_with_macro_and_complex_password.xls";
    let (temp_dir, temp_file) = create_temp_dir(&file, 6);
    xl_97(Path::new(&temp_file), true).unwrap();
    let (p, _) = read::xl_97_project(&temp_file, false).unwrap();
    assert!(!p.is_locked());
    let _ = std::fs::remove_dir_all(temp_dir);
}

fn replacement_filename(source: &dyn AsRef<Path>) -> PathBuf {
    let source = source.as_ref();
    let mut new = PathBuf::from(source);
    let mut stem = source.file_stem().unwrap().to_owned();
    stem.push("_unlocked");
    new.set_file_name(stem);
    let ext = source.extension().unwrap();
    new.set_extension(ext);
    new
}

fn create_temp_dir(source: &dyn AsRef<Path>, unique_num: u8) -> (PathBuf, PathBuf) {
    let source = source.as_ref();
    let mut folder = source.parent().unwrap().to_path_buf();
    folder.push(format!("temp_{unique_num}"));
    let mut copied_file = folder.clone();
    copied_file.push(source.file_name().unwrap());
    std::fs::create_dir(&folder).unwrap();
    let _ = std::fs::copy(source, &copied_file);
    (folder, copied_file)
}
