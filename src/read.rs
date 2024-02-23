use std::fs::File;
use std::io::{BufRead, Cursor, Read};
use std::path::Path;

use crate::consts;
use crate::error::UnlockError;
use crate::error::UnlockResult;
use cfb::Stream;
use vba_password::ProjectPassword;
use vba_protection_state::ProjectProtectionState;
use vba_visibility::ProjectVisibililyState;
use zip::ZipArchive;

pub fn xl_97(filename: &Path, decode: bool) -> UnlockResult<()> {
    let mut file = cfb::open(filename).map_err(UnlockError::CFBOpen)?;
    let project = file.open_stream(consts::CFB_VBA_PATH)?;
    print_info(project, decode)
}

pub fn xl(filename: &Path, decode: bool) -> UnlockResult<()> {
    let zipfile = File::open(filename)?;
    let mut archive = zip::ZipArchive::new(zipfile)?;
    let vba_raw = zip_to_raw_vba(&mut archive)?;
    let mut vba_cfb = cfb::CompoundFile::open(vba_raw).map_err(UnlockError::CFBOpen)?;
    let project = vba_cfb.open_stream(consts::PROJECT_PATH)?;
    print_info(project, decode)
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

fn print_info<T: std::io::Read + std::io::Seek>(
    project: Stream<T>,
    decode: bool,
) -> UnlockResult<()> {
    for line in project.lines().map_while(Result::ok) {
        if line.starts_with("CMG=") {
            let protection_state: ProjectProtectionState = line.parse()?;
            println!("{protection_state}");
        }
        if line.starts_with("DPB=") {
            let password: ProjectPassword = line.parse()?;
            print!("{password}");
            if password.is_hash() && decode {
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
    use crate::error::ProtectionState;
    use crate::ovba::algorithms::data_encryption;
    use std::{fmt::Display, str::FromStr};

    pub struct ProjectProtectionState {
        user: bool,
        host: bool,
        vbe: bool,
    }

    impl FromStr for ProjectProtectionState {
        type Err = ProtectionState;

        fn from_str(s: &str) -> Result<Self, Self::Err> {
            let hex_str = s.trim_start_matches("CMG=\"").trim_end_matches('"');
            let data = data_encryption::decode_str(hex_str)?.into_inner();
            if data.len() != 4 {
                return Err(ProtectionState::DataLength(data.len()));
            }
            if data[0] > 7 || data[1] != 0 || data[2] != 0 || data[3] != 0 {
                return Err(ProtectionState::ReservedBits([
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

mod vba_password {
    // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/79685426-30fe-43cd-9cbf-7f161c3de7d8
    use crate::ovba::algorithms::data_encryption;
    use crate::{error, ovba::algorithms::password_hash};
    use sha1::{Digest, Sha1};
    use std::{fmt::Display, str::FromStr};

    pub enum ProjectPassword {
        None,
        Hash([u8; 4], [u8; 20]),
        PlainText(String),
    }

    impl FromStr for ProjectPassword {
        type Err = error::Password;

        fn from_str(s: &str) -> Result<Self, Self::Err> {
            let hex_str = s.trim_start_matches("DPB=\"").trim_end_matches('"');
            let data = data_encryption::decode_str(hex_str)?.into_inner();
            Ok(match data.len() {
                0 => return Err(error::Password::NoData),
                1 => Self::new_none(&data)?,
                29 => {
                    let (salt, hash) = password_hash::decode(data)?;
                    Self::Hash(salt, hash)
                }
                _ => Self::new_plain(&data)?,
            })
        }
    }

    impl ProjectPassword {
        pub const fn is_hash(&self) -> bool {
            matches!(self, Self::Hash(_, _))
        }

        fn new_none(data: &[u8]) -> Result<Self, error::PasswordNone> {
            if data.first() != Some(0x00).as_ref() {
                return Err(error::PasswordNone::NotNull(data[0]));
            }
            Ok(Self::None)
        }

        fn new_plain(data: &[u8]) -> Result<Self, error::PasswordPlain> {
            if data.last() != Some(0x00).as_ref() {
                return Err(error::PasswordPlain::Terminator(
                    *data
                        .last()
                        .expect("Cannot construct a plain password with zero length data"),
                ));
            }
            let password = String::from_utf8_lossy(&data[0..(data.len() - 1)]).to_string();
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

mod vba_visibility {
    // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/690c96e8-e862-497f-bb7d-5eacf4dc742a
    use crate::error;
    use crate::ovba::algorithms::data_encryption;
    use std::{fmt::Display, str::FromStr};

    pub enum ProjectVisibililyState {
        NotVisible,
        Visible,
    }

    impl FromStr for ProjectVisibililyState {
        type Err = error::Visibility;

        fn from_str(s: &str) -> Result<Self, Self::Err> {
            let hex_str = s.trim_start_matches("GC=\"").trim_end_matches('"');
            let data = data_encryption::decode_str(hex_str)?.into_inner();
            if data.len() != 1 {
                return Err(error::Visibility::DataLength(data.len()));
            }
            match data.first() {
                Some(0x00) => Ok(Self::NotVisible),
                Some(0xff) => Ok(Self::Visible),
                Some(x) => Err(error::Visibility::InvalidState(*x)),
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
