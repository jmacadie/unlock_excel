use std::path::Path;
use unlock_excel::read::{xl_97_project, xl_project};

/*
* XLSM
* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

#[test]
fn read_unlocked_no_decode_xlsm() {
    let (p, d) = xl_project(Path::new("tests/data/xlsm/Unlocked_with_macro.xlsm"), false).unwrap();
    assert!(!p.is_locked());
    assert!(d.is_none());
}

#[test]
fn read_locked_1_no_decode_xlsm() {
    let (p, d) = xl_project(Path::new("tests/data/xlsm/Locked_with_macro.xlsm"), false).unwrap();
    assert!(p.is_locked());
    assert!(d.is_none());
}

#[test]
fn read_locked_2_no_decode_xlsm() {
    let (p, d) = xl_project(
        Path::new("tests/data/xlsm/Locked_with_macro_and_complex_password.xlsm"),
        false,
    )
    .unwrap();
    assert!(p.is_locked());
    assert!(d.is_none());
}

#[test]
fn read_unlocked_decode_xlsm() {
    let (p, d) = xl_project(Path::new("tests/data/xlsm/Unlocked_with_macro.xlsm"), true).unwrap();
    assert!(!p.is_locked());
    assert!(d.is_none());
}

#[test]
fn read_locked_1_decode_xlsm() {
    let (p, d) = xl_project(Path::new("tests/data/xlsm/Locked_with_macro.xlsm"), true).unwrap();
    assert!(p.is_locked());
    assert_eq!(Some("P@ssw0rd"), d.as_deref());
}

#[test]
fn read_locked_2_decode_xlsm() {
    let (p, d) = xl_project(
        Path::new("tests/data/xlsm/Locked_with_macro_and_complex_password.xlsm"),
        true,
    )
    .unwrap();
    assert!(p.is_locked());
    assert!(d.is_none());
}

/*
* XLSB
* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

#[test]
fn read_unlocked_no_decode_xlsb() {
    let (p, d) = xl_project(Path::new("tests/data/xlsb/Unlocked_with_macro.xlsb"), false).unwrap();
    assert!(!p.is_locked());
    assert!(d.is_none());
}

#[test]
fn read_locked_1_no_decode_xlsb() {
    let (p, d) = xl_project(Path::new("tests/data/xlsb/Locked_with_macro.xlsb"), false).unwrap();
    assert!(p.is_locked());
    assert!(d.is_none());
}

#[test]
fn read_locked_2_no_decode_xlsb() {
    let (p, d) = xl_project(
        Path::new("tests/data/xlsb/Locked_with_macro_and_complex_password.xlsb"),
        false,
    )
    .unwrap();
    assert!(p.is_locked());
    assert!(d.is_none());
}

#[test]
fn read_unlocked_decode_xlsb() {
    let (p, d) = xl_project(Path::new("tests/data/xlsb/Unlocked_with_macro.xlsb"), true).unwrap();
    assert!(!p.is_locked());
    assert!(d.is_none());
}

#[test]
fn read_locked_1_decode_xlsb() {
    let (p, d) = xl_project(Path::new("tests/data/xlsb/Locked_with_macro.xlsb"), true).unwrap();
    assert!(p.is_locked());
    assert_eq!(Some("P@ssw0rd"), d.as_deref());
}

#[test]
fn read_locked_2_decode_xlsb() {
    let (p, d) = xl_project(
        Path::new("tests/data/xlsb/Locked_with_macro_and_complex_password.xlsb"),
        true,
    )
    .unwrap();
    assert!(p.is_locked());
    assert!(d.is_none());
}

/*
* XLS
* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

#[test]
fn read_unlocked_no_decode_xls() {
    let (p, d) = xl_97_project(Path::new("tests/data/xls/Unlocked_with_macro.xls"), false).unwrap();
    assert!(!p.is_locked());
    assert!(d.is_none());
}

#[test]
fn read_locked_1_no_decode_xls() {
    let (p, d) = xl_97_project(Path::new("tests/data/xls/Locked_with_macro.xls"), false).unwrap();
    assert!(p.is_locked());
    assert!(d.is_none());
}

#[test]
fn read_locked_2_no_decode_xls() {
    let (p, d) = xl_97_project(
        Path::new("tests/data/xls/Locked_with_macro_and_complex_password.xls"),
        false,
    )
    .unwrap();
    assert!(p.is_locked());
    assert!(d.is_none());
}

#[test]
fn read_unlocked_decode_xls() {
    let (p, d) = xl_97_project(Path::new("tests/data/xls/Unlocked_with_macro.xls"), true).unwrap();
    assert!(!p.is_locked());
    assert!(d.is_none());
}

#[test]
fn read_locked_1_decode_xls() {
    let (p, d) = xl_97_project(Path::new("tests/data/xls/Locked_with_macro.xls"), true).unwrap();
    assert!(p.is_locked());
    assert_eq!(Some("P@ssw0rd"), d.as_deref());
}

#[test]
fn read_locked_2_decode_xls() {
    let (p, d) = xl_97_project(
        Path::new("tests/data/xls/Locked_with_macro_and_complex_password.xls"),
        true,
    )
    .unwrap();
    assert!(p.is_locked());
    assert!(d.is_none());
}
