use std::io::BufRead;

use crate::error;
use crate::error::UnlockResult;
use cfb::Stream;
use vba_password::ProjectPassword;
use vba_protection_state::ProjectProtectionState;
use vba_visibility::ProjectVisibililyState;

pub fn print_info<T: std::io::Read + std::io::Seek>(
    project: Stream<T>,
    decode: bool,
) -> UnlockResult<()> {
    for line in project.lines().flatten() {
        if line.starts_with("CMG=") {
            let protection_state: ProjectProtectionState = line.parse()?;
            println!("{protection_state}");
        }
        if line.starts_with("DPB=") {
            let password: ProjectPassword = line.parse()?;
            print!("{password}");
            if decode {
                password.crack_password().map_or_else(|| {
                    println!("  Was unable to decode the password. Try removing the password, which always works");
                }, |clear| {
                    println!("  Decoded Password: {clear}");
                });
            }
            println!();
        }
        if line.starts_with("GC=") {
            let visibility_state: ProjectVisibililyState = line.parse()?;
            println!("{visibility_state}");
        }
    }
    Ok(())
}

mod vba_protection_state {
    // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/55e770e2-e1a4-4d1c-a8a4-dcfca27d6663
    use super::decrypt;
    use crate::error::VBAProtectionState;
    use std::{fmt::Display, str::FromStr};

    pub struct ProjectProtectionState {
        user: bool,
        host: bool,
        vbe: bool,
    }

    impl FromStr for ProjectProtectionState {
        type Err = VBAProtectionState;

        fn from_str(s: &str) -> Result<Self, Self::Err> {
            let hex_str = s.trim_start_matches("CMG=\"").trim_end_matches('"');
            let (data, length) = decrypt(hex_str)?;
            if length != 4 {
                return Err(VBAProtectionState::Length(length));
            }
            if data.len() != 4 {
                return Err(VBAProtectionState::DataLength(data.len()));
            }
            if data[0] > 7 || data[1] != 0 || data[2] != 0 || data[3] != 0 {
                return Err(VBAProtectionState::ReservedBits([
                    data[0], data[1], data[2], data[3],
                ]));
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
}

pub mod vba_password {
    // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/79685426-30fe-43cd-9cbf-7f161c3de7d8
    use super::decrypt;
    use crate::error;
    use sha1::{Digest, Sha1};
    use std::{fmt::Display, str::FromStr};

    pub enum ProjectPassword {
        None,
        Hash([u8; 4], [u8; 20]),
        PlainText(String),
    }

    impl FromStr for ProjectPassword {
        type Err = error::VBAPassword;

        fn from_str(s: &str) -> Result<Self, Self::Err> {
            let hex_str = s.trim_start_matches("DPB=\"").trim_end_matches('"');
            let (data, length) = decrypt(hex_str)?;
            Ok(match length {
                1 => Self::new_none(&data)?,
                29 => Self::new_hash(&data)?,
                l => Self::new_plain(&data, l)?,
            })
        }
    }

    impl ProjectPassword {
        fn new_none(data: &[u8]) -> Result<Self, error::VBAPasswordNone> {
            if data.len() != 1 {
                return Err(error::VBAPasswordNone::DataLength(data.len()));
            }
            if data.first() != Some(0x00).as_ref() {
                return Err(error::VBAPasswordNone::NotNull(data[0]));
            }
            Ok(Self::None)
        }

        fn new_hash(data: &[u8]) -> Result<Self, error::VBAPasswordHash> {
            // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/9d9f81e6-f92e-4338-a242-d38c1fcceed6
            if data.len() != 29 {
                return Err(error::VBAPasswordHash::DataLength(data.len()));
            }
            if data.first() != Some(0xff).as_ref() {
                return Err(error::VBAPasswordHash::Reserved(data[0]));
            }
            if data.last() != Some(0x00).as_ref() {
                return Err(error::VBAPasswordHash::Terminator(data[28]));
            }

            let mut salt = [0; 4];
            salt.clone_from_slice(&data[4..8]);

            let mut hash = [0; 20];
            hash.clone_from_slice(&data[8..28]);

            // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/5797c2e1-4c86-4f44-89b4-1edb30da00cc
            // Add nulls to salt
            let mut grbitkey = data[1] & 0x0f;
            for (i, byte) in salt.iter_mut().enumerate() {
                if grbitkey & 1 == 0 {
                    if *byte != 0x01 {
                        return Err(error::VBAPasswordHash::SaltNull(salt, i));
                    }
                    *byte = 0;
                }
                grbitkey >>= 1;
            }

            // Add nulls to hash
            let mut grbithashnull = u32::from(data[1]) >> 4;
            grbithashnull |= u32::from(data[2]) << 4;
            grbithashnull |= u32::from(data[3]) << 12;
            for (i, byte) in hash.iter_mut().enumerate() {
                if grbithashnull & 1 == 0 {
                    if *byte != 0x01 {
                        return Err(error::VBAPasswordHash::HashNull(hash, i));
                    }
                    *byte = 0;
                }
                grbithashnull >>= 1;
            }

            Ok(Self::Hash(salt, hash))
        }

        fn new_plain(data: &[u8], length: u32) -> Result<Self, error::VBAPasswordPlain> {
            let Ok(length) = usize::try_from(length) else {
                return Err(error::VBAPasswordPlain::LengthToUsize(length));
            };
            if data.len() != length {
                return Err(error::VBAPasswordPlain::DataLength(data.len(), length));
            }
            if data.last() != Some(0x00).as_ref() {
                return Err(error::VBAPasswordPlain::Terminator(data[length - 1]));
            }
            let password = String::from_utf8_lossy(&data[0..(length - 2)]).to_string();
            Ok(Self::PlainText(password))
        }

        pub fn crack_password(&self) -> Option<String> {
            if let Self::Hash(salt, hash) = self {
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
            }
            None
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
                    writeln!(f)?;
                }
                Self::PlainText(password) => write!(f, "{password} (plain-text)")?,
            }
            Ok(())
        }
    }
}

pub mod vba_visibility {
    // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/690c96e8-e862-497f-bb7d-5eacf4dc742a
    use super::decrypt;
    use crate::error;
    use std::{fmt::Display, str::FromStr};

    pub enum ProjectVisibililyState {
        NotVisible,
        Visible,
    }

    impl FromStr for ProjectVisibililyState {
        type Err = error::VBAVisibility;

        fn from_str(s: &str) -> Result<Self, Self::Err> {
            let hex_str = s.trim_start_matches("GC=\"").trim_end_matches('"');
            let (data, length) = decrypt(hex_str)?;
            if length != 1 {
                return Err(error::VBAVisibility::Length(length));
            }
            if data.len() != 1 {
                return Err(error::VBAVisibility::DataLength(data.len()));
            }
            match data.first() {
                Some(0x00) => Ok(Self::NotVisible),
                Some(0xff) => Ok(Self::Visible),
                Some(x) => Err(error::VBAVisibility::InvalidState(*x)),
                None => unreachable!(),
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
}

fn decrypt(hex: &str) -> Result<(Vec<u8>, u32), error::VBADecrypt> {
    // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/a2ad3aa7-e180-4ccb-8511-7e0eb49a0ad9
    // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/7e9d84fe-86e3-46d6-aaff-8388e72c0168

    if hex.bytes().any(|x| !x.is_ascii_hexdigit()) {
        return Err(error::VBADecrypt::InvalidHex(hex.to_owned()));
    }

    let seed = u8::from_str_radix(&hex[0..2], 16)?;
    let version_enc = u8::from_str_radix(&hex[2..4], 16)?;
    let project_key_enc = u8::from_str_radix(&hex[4..6], 16)?;

    let version = seed ^ version_enc;
    if version != 2 {
        return Err(error::VBADecrypt::Version(version));
    }
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

    Ok((data, length))
}
