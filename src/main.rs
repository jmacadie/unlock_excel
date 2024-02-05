#![warn(clippy::all, clippy::pedantic, clippy::nursery)]

use cfb::CompoundFile;
use std::{fmt::Display, io::prelude::*, str::FromStr};

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
    let fname = std::path::Path::new(&*args[1]);
    let zipfile = std::fs::File::open(fname).unwrap();

    let mut archive = zip::ZipArchive::new(zipfile).unwrap();

    let Ok(mut file) = archive.by_name("xl/vbaProject.bin") else {
        eprintln!("VBA section not found in file");
        return 2;
    };

    // Read the uncompressed bytes of the vbaProject.bin file into an in-memory cursor
    let mut buffer = Vec::with_capacity(1024);
    let _ = file.read_to_end(&mut buffer);
    let raw = std::io::Cursor::new(buffer);

    let mut cfb_file = match CompoundFile::open(raw) {
        Ok(f) => f,
        Err(e) => {
            eprintln!("Error in the structure of the vbaProject.bin CFB file: {e}");
            return 3;
        }
    };

    let project = match cfb_file.open_stream(PROJECT_PATH) {
        Ok(s) => s,
        Err(e) => {
            eprintln!("Cannot open the Project Stream: {e}");
            return 4;
        }
    };

    for line in project.lines().flatten() {
        if line.starts_with("CMG=") {
            let protection_state: ProjectProtectionState = line.parse().unwrap();
            println!("{protection_state}");
        }
        if line.starts_with("DPB=") {
            let password: ProjectPassword = line.parse().unwrap();
            println!("{password}");
        }
        if line.starts_with("GC=") {
            let visibility_state: ProjectVisibililyState = line.parse().unwrap();
            println!("{visibility_state}");
        }
    }

    0
}

// https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/55e770e2-e1a4-4d1c-a8a4-dcfca27d6663
struct ProjectProtectionState {
    user: bool,
    host: bool,
    vbe: bool,
}

impl FromStr for ProjectProtectionState {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        let hex_str = s.trim_start_matches("CMG=\"").trim_end_matches('"');
        let (data, length) = decrypt(hex_str);
        if length != 4 {
            return Err(format!(
                "Length of Project Protection State is expected to be 4, not {length}"
            ));
        }
        if data[0] > 7 || data[1] != 0 || data[2] != 0 || data[3] != 0 {
            return Err(format!(
                "Reserved bits (above first 3) should all be zero: {data:?}"
            ));
        }
        let user_protected = data[0] & 1 == 1;
        let host_protected = data[0] & 2 == 2;
        let vbe_protected = data[0] & 4 == 4;
        Ok(Self {
            user: user_protected,
            host: host_protected,
            vbe: vbe_protected,
        })
    }
}

impl Display for ProjectProtectionState {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        writeln!(f, "Project Protection State:")?;
        writeln!(f, "  User Protected: {}", self.user)?;
        writeln!(f, "  Host Protected: {}", self.host)?;
        writeln!(f, "  VBE Protected: {}", self.vbe)?;
        Ok(())
    }
}

// https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/79685426-30fe-43cd-9cbf-7f161c3de7d8
enum ProjectPassword {
    None,
    Hash([u8; 4], [u8; 20]),
    PlainText(String),
}

impl FromStr for ProjectPassword {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        let hex_str = s.trim_start_matches("DPB=\"").trim_end_matches('"');
        let (data, length) = decrypt(hex_str);
        match length {
            1 => {
                if data.first() != Some(0x00).as_ref() {
                    return Err(format!(
                        "A VBA project with no password should have data value of 0: {}",
                        data[0]
                    ));
                }
                Ok(Self::None)
            }
            29 => {
                // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/9d9f81e6-f92e-4338-a242-d38c1fcceed6
                if data.len() != 29 {
                    return Err(format!(
                        "Hashed data is meant to be 29 bytes long, not {}",
                        data.len()
                    ));
                }
                if data.first() != Some(0xff).as_ref() {
                    return Err(format!("The first byte of a hashed password it reserved and MUST be 0xff, not 0x{:02x}", data[0]));
                }
                if data.last() != Some(0x00).as_ref() {
                    return Err(format!(
                        "The terminator byte of the hashed password MUST be 0x00, not 0x{:02x}",
                        data[28]
                    ));
                }

                let mut salt = [0; 4];
                salt.clone_from_slice(&data[4..8]);

                let mut sha1 = [0; 20];
                sha1.clone_from_slice(&data[8..28]);

                // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/5797c2e1-4c86-4f44-89b4-1edb30da00cc
                // Add nulls to salt
                let mut grbitkey = data[1] & 0x0f;
                for byte in &mut salt {
                    if grbitkey & 1 == 0 {
                        if *byte != 0x01 {
                            return Err(format!("Replacing byte in the salt with a null. It was expected to be 0x01 before replacement, not 0x{byte:02x}"));
                        }
                        *byte = 0;
                    }
                    grbitkey >>= 1;
                }

                // Add nulls to hash
                let mut grbithashnull = u32::from(data[1]) >> 4;
                grbithashnull |= u32::from(data[2]) << 4;
                grbithashnull |= u32::from(data[3]) << 12;
                for byte in &mut sha1 {
                    if grbithashnull & 1 == 0 {
                        if *byte != 0x01 {
                            return Err(format!("Replacing byte in the hash with a null. It was expected to be 0x01 before replacement, not 0x{byte:02x}"));
                        }
                        *byte = 0;
                    }
                    grbithashnull >>= 1;
                }

                Ok(Self::Hash(salt, sha1))
            }
            _ => {
                let Ok(length) = usize::try_from(length) else {
                    return Err(format!("Cannot convert {length}"));
                };
                if data.len() != length {
                    return Err(format!(
                        "Plain text password data is meant to be {length} bytes long, not {}",
                        data.len()
                    ));
                }
                if data.last() != Some(0x00).as_ref() {
                    return Err(format!(
                        "The terminator byte of the plain text password MUST be 0x00, not 0x{:02x}",
                        data[length - 1]
                    ));
                }
                let password = String::from_utf8_lossy(&data[0..(length - 2)]).to_string();
                Ok(Self::PlainText(password))
            }
        }
    }
}

impl Display for ProjectPassword {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Project Password: ")?;
        match self {
            Self::None => write!(f, "None")?,
            Self::Hash(salt, hash) => {
                writeln!(f, "Hashed (SHA1)")?;
                write!(f, "  Salt: ")?;
                for byte in salt {
                    write!(f, "{byte:02x}")?;
                }
                writeln!(f)?;
                write!(f, "  SHA1 Hash: ")?;
                for byte in hash {
                    write!(f, "{byte:02x}")?;
                }
            }
            Self::PlainText(password) => write!(f, "{password} (plain-text)")?,
        }
        Ok(())
    }
}

// https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/690c96e8-e862-497f-bb7d-5eacf4dc742a
enum ProjectVisibililyState {
    NotVisible,
    Visible,
}

impl FromStr for ProjectVisibililyState {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        let hex_str = s.trim_start_matches("GC=\"").trim_end_matches('"');
        let (data, length) = decrypt(hex_str);
        if length != 1 {
            return Err(format!(
                "Length of Project Protection State is expected to be 1, not {length}"
            ));
        }
        match data.first() {
            Some(0x00) => Ok(Self::NotVisible),
            Some(0xff) => Ok(Self::Visible),
            _ => Err(format!("Invalid value for project visibility: {}", data[0])),
        }
    }
}

impl Display for ProjectVisibililyState {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        writeln!(f, "Project Visibility:")?;
        match self {
            Self::Visible => writeln!(f, "  Visible")?,
            Self::NotVisible => writeln!(f, "  Not Visible")?,
        }
        Ok(())
    }
}

fn decrypt(hex: &str) -> (Vec<u8>, u32) {
    // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/a2ad3aa7-e180-4ccb-8511-7e0eb49a0ad9
    // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/7e9d84fe-86e3-46d6-aaff-8388e72c0168

    let seed = u8::from_str_radix(&hex[0..2], 16).unwrap();
    let version_enc = u8::from_str_radix(&hex[2..4], 16).unwrap();
    let project_key_enc = u8::from_str_radix(&hex[4..6], 16).unwrap();

    let version = seed ^ version_enc;
    assert!(version == 2, "Version MUST be 2");
    let project_key = seed ^ project_key_enc;
    let ignored_length = ((seed & 6) >> 1).into();

    let mut unencrypted_byte_1 = project_key;
    let mut encrypted_byte_1 = project_key_enc;
    let mut encrypted_byte_2 = version_enc;

    // Convert remaining string into a u8 (byte) iterator
    let data_bytes = hex[6..].as_bytes().chunks_exact(2).map(|x| {
        let upper = match x[0] {
            val if val.is_ascii_digit() => val - b'0',
            val if val.is_ascii_lowercase() => val - b'a' + 10,
            val if val.is_ascii_uppercase() => val - b'A' + 10,
            _ => unreachable!(),
        };
        let lower = match x[1] {
            val if val.is_ascii_digit() => val - b'0',
            val if val.is_ascii_lowercase() => val - b'a' + 10,
            val if val.is_ascii_uppercase() => val - b'A' + 10,
            _ => unreachable!(),
        };
        (upper << 4) | lower
    });

    // Generate the length & data
    let mut data = Vec::new();
    let mut length = 0;
    for (i, byte_enc) in data_bytes.enumerate() {
        let byte = byte_enc ^ (encrypted_byte_2 + unencrypted_byte_1);
        encrypted_byte_2 = encrypted_byte_1;
        encrypted_byte_1 = byte_enc;
        unencrypted_byte_1 = byte;
        match i {
            x if x < ignored_length => (), // Ignore these bytes
            x if x < ignored_length + 4 => {
                let b = u32::from(byte);
                let shift = 4 * (x - ignored_length);
                length |= b << shift;
            }
            _ => data.push(byte),
        }
    }

    (data, length)
}
