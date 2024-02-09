use std::{
    fmt::{Debug, Display},
    io,
    num::ParseIntError,
};

pub type UnlockResult<T> = Result<T, UnlockError>;

#[allow(clippy::module_name_repetitions)]
pub enum UnlockError {
    FileOpen(io::Error),
    Zip(zip::result::ZipError),
    NoVBAFile,
    CFBOpen(io::Error),
    ProjectStructure(VBAProjectStructure),
}

impl From<io::Error> for UnlockError {
    fn from(value: io::Error) -> Self {
        Self::FileOpen(value)
    }
}

impl From<zip::result::ZipError> for UnlockError {
    fn from(value: zip::result::ZipError) -> Self {
        Self::Zip(value)
    }
}

impl From<VBAProtectionState> for UnlockError {
    fn from(value: VBAProtectionState) -> Self {
        Self::ProjectStructure(VBAProjectStructure::ProtectionState(value))
    }
}

impl From<VBAPassword> for UnlockError {
    fn from(value: VBAPassword) -> Self {
        Self::ProjectStructure(VBAProjectStructure::Password(value))
    }
}

impl From<VBAVisibility> for UnlockError {
    fn from(value: VBAVisibility) -> Self {
        Self::ProjectStructure(VBAProjectStructure::Visibility(value))
    }
}

impl Display for UnlockError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::FileOpen(e) => write!(f, "{e}"),
            Self::Zip(e) => write!(
                f,
                "Problem with the zip representation of the supplied Excel file: {e}"
            ),
            Self::NoVBAFile => write!(
                f,
                "Could not find the 'xl/vbaProject.bin' file within the extracted archive"
            ),
            Self::CFBOpen(e) => write!(
                f,
                "There was a problem reading the CFB format vbaProject.bin file: {e}"
            ),
            Self::ProjectStructure(e) => write!(f, "{e}"),
        }
    }
}

impl Debug for UnlockError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{self}")
    }
}

pub enum VBAProjectStructure {
    ProtectionState(VBAProtectionState),
    Password(VBAPassword),
    Visibility(VBAVisibility),
}

impl Display for VBAProjectStructure {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::ProtectionState(e) => write!(f, "{e}"),
            Self::Password(e) => write!(f, "{e}"),
            Self::Visibility(e) => write!(f, "{e}"),
        }
    }
}

impl From<VBAProtectionState> for VBAProjectStructure {
    fn from(value: VBAProtectionState) -> Self {
        Self::ProtectionState(value)
    }
}

impl From<VBAPassword> for VBAProjectStructure {
    fn from(value: VBAPassword) -> Self {
        Self::Password(value)
    }
}

impl From<VBAVisibility> for VBAProjectStructure {
    fn from(value: VBAVisibility) -> Self {
        Self::Visibility(value)
    }
}

#[derive(Debug)]
pub enum VBAProtectionState {
    Decrypt(VBADecrypt),
    Length(u32),
    DataLength(usize),
    ReservedBits([u8; 4]),
}

impl Display for VBAProtectionState {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Decrypt(e) => write!(f, "{e}"),
            Self::Length(l) => write!(f, "The length parameter for Data Encryption of the VBA Project Protection State MUST be 4, not {l}"),
            Self::DataLength(l) => write!(f, "The Data array for Data Encryption of the VBA Project Protection State SHOULD be 4 bytes, not {l} bytes"),
            Self::ReservedBits(data) => write!(f, "The upper 29 bits of data are reserved and MUST all be 0. Data decodeded to {:08b}{:08b}{:08b}{:08b}", data[0], data[1], data[2], data[3]),
        }
    }
}

impl From<VBADecrypt> for VBAProtectionState {
    fn from(value: VBADecrypt) -> Self {
        Self::Decrypt(value)
    }
}

#[derive(Debug)]
pub enum VBAPassword {
    Decrypt(VBADecrypt),
    None(VBAPasswordNone),
    Hash(VBAPasswordHash),
    PlainText(VBAPasswordPlain),
}

impl Display for VBAPassword {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Decrypt(e) => write!(f, "{e}"),
            Self::None(e) => write!(f, "{e}"),
            Self::Hash(e) => write!(f, "{e}"),
            Self::PlainText(e) => write!(f, "{e}"),
        }
    }
}

impl From<VBADecrypt> for VBAPassword {
    fn from(value: VBADecrypt) -> Self {
        Self::Decrypt(value)
    }
}

impl From<VBAPasswordNone> for VBAPassword {
    fn from(value: VBAPasswordNone) -> Self {
        Self::None(value)
    }
}

impl From<VBAPasswordHash> for VBAPassword {
    fn from(value: VBAPasswordHash) -> Self {
        Self::Hash(value)
    }
}

impl From<VBAPasswordPlain> for VBAPassword {
    fn from(value: VBAPasswordPlain) -> Self {
        Self::PlainText(value)
    }
}

#[derive(Debug)]
pub enum VBAPasswordNone {
    DataLength(usize),
    NotNull(u8),
}

impl Display for VBAPasswordNone {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::DataLength(l) => write!(
                f,
                "The data length for a VBA project without a password MUST be 1, not {l}"
            ),
            Self::NotNull(b) => write!(
                f,
                "The data value for a VBA project without a password MUST be 0x00, not 0x{b:02x}"
            ),
        }
    }
}

#[derive(Debug)]
pub enum VBAPasswordHash {
    DataLength(usize),
    Reserved(u8),
    Terminator(u8),
    SaltNull([u8; 4], usize),
    HashNull([u8; 20], usize),
}

impl Display for VBAPasswordHash {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::DataLength(l) => write!(f, "The length of the VBA password hash data structure MUST be 29, not {l}"),
            Self::Reserved(b) => write!(f, "The first byte of the VBA password hash data structure is reserved and MUST be 0xff, not 0x{b:02x}"),
            Self::Terminator(b) => write!(f, "The final byte of the VBA password hash data structure is the terminator and MUST be 0x00, not 0x{b:02x}"),
            Self::SaltNull(data, i) => write!(f, "The byte in position {i} of the salt {data:?} is being replaced with a null. It should have a value of 1 before update"),
            Self::HashNull(data, i) => write!(f, "The byte in position {i} of the hash {data:?} is being replaced with a null. It should have a value of 1 before update"),
        }
    }
}

#[derive(Debug)]
pub enum VBAPasswordPlain {
    LengthToUsize(u32),
    DataLength(usize, usize),
    Terminator(u8),
}

impl Display for VBAPasswordPlain {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::LengthToUsize(_) => write!(f, "Could not convert the length to a usize (to index into the data array). This should never happen!"),
            Self::DataLength(dl, l) => write!(f, "The length of the decrypted plain-text password {dl}, does not match the decrypted length {l}. Values cannot be trusted"),
            Self::Terminator(b) => write!(f, "The plain-text password MUST be null terminated. We got 0x{b:02x}"),
        }
    }
}

#[derive(Debug)]
pub enum VBAVisibility {
    Decrypt(VBADecrypt),
    Length(u32),
    DataLength(usize),
    InvalidState(u8),
}

impl Display for VBAVisibility {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Decrypt(e) => write!(f, "{e}"),
            Self::Length(l) => write!(f, "VBA Visibility should decrypt to a length of 1, not {l}"),
            Self::DataLength(l) => write!(
                f,
                "VBA Visibility should decrypt to a data length of 1, not {l}"
            ),
            Self::InvalidState(s) => write!(
                f,
                "VBA Visibility only has two valid values: 0x00 and 0xff. 0x{s:02x} was found"
            ),
        }
    }
}

impl From<VBADecrypt> for VBAVisibility {
    fn from(value: VBADecrypt) -> Self {
        Self::Decrypt(value)
    }
}

#[derive(Debug)]
pub enum VBADecrypt {
    InvalidHex(String),
    Version(u8),
}

impl From<ParseIntError> for VBADecrypt {
    fn from(value: ParseIntError) -> Self {
        Self::InvalidHex(format!("{value}"))
    }
}

impl Display for VBADecrypt {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::InvalidHex(s) => write!(
                f,
                "Cannot apply VBA data decryption as supplied value is not valid hex: {s}"
            ),
            Self::Version(v) => {
                writeln!(f, "VBA Data Encryption Version MUST be 2 according to the spec: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/7e9d84fe-86e3-46d6-aaff-8388e72c0168")?;
                writeln!(f, "Got {v}")
            }
        }
    }
}
