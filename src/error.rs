use std::{
    fmt::{Debug, Display},
    io,
    num::ParseIntError,
};

pub type UnlockResult<T> = Result<T, UnlockError>;

#[allow(clippy::module_name_repetitions)]
pub enum UnlockError {
    FileOpen(io::Error),
    NotExcel(String),
    XlsX(String),
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
            Self::NotExcel(file) => write!(f, "{file} is not an Excel file. Try harder"),
            Self::XlsX(file) => write!(
                f,
                "{file} is Excel's format for files with no VBA. There is nothing to operate on"
            ),
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

#[derive(Debug, PartialEq, Eq)]
pub enum VBAProtectionState {
    Decrypt(VBADecrypt),
    DataLength(usize),
    ReservedBits([u8; 4]),
}

impl Display for VBAProtectionState {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Decrypt(e) => write!(f, "{e}"),
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

#[derive(Debug, PartialEq, Eq)]
pub enum VBAPassword {
    Decrypt(VBADecrypt),
    None(VBAPasswordNone),
    Hash(VBAPasswordHash),
    PlainText(VBAPasswordPlain),
    NoData,
}

impl Display for VBAPassword {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Decrypt(e) => write!(f, "{e}"),
            Self::None(e) => write!(f, "{e}"),
            Self::Hash(e) => write!(f, "{e}"),
            Self::PlainText(e) => write!(f, "{e}"),
            Self::NoData => write!(f, "The decrypted password had no data"),
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

#[derive(Debug, PartialEq, Eq)]
pub enum VBAPasswordNone {
    NotNull(u8),
}

impl Display for VBAPasswordNone {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::NotNull(b) => write!(
                f,
                "The data value for a VBA project without a password MUST be 0x00, not 0x{b:02x}"
            ),
        }
    }
}

#[derive(Debug, PartialEq, Eq)]
pub enum VBAPasswordHash {
    Reserved(u8),
    Terminator(u8),
    SaltNull([u8; 4], usize),
    HashNull([u8; 20], usize),
}

impl Display for VBAPasswordHash {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Reserved(b) => write!(f, "The first byte of the VBA password hash data structure is reserved and MUST be 0xff, not 0x{b:02x}"),
            Self::Terminator(b) => write!(f, "The final byte of the VBA password hash data structure is the terminator and MUST be 0x00, not 0x{b:02x}"),
            Self::SaltNull(data, i) => write!(f, "The byte in position {i} of the salt {data:?} is being replaced with a null. It should have a value of 1 before update"),
            Self::HashNull(data, i) => write!(f, "The byte in position {i} of the hash {data:?} is being replaced with a null. It should have a value of 1 before update"),
        }
    }
}

#[derive(Debug, PartialEq, Eq)]
pub enum VBAPasswordPlain {
    Terminator(u8),
}

impl Display for VBAPasswordPlain {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Terminator(b) => write!(
                f,
                "The plain-text password MUST be null terminated. We got 0x{b:02x}"
            ),
        }
    }
}

#[derive(Debug, PartialEq, Eq)]
pub enum VBAVisibility {
    Decrypt(VBADecrypt),
    DataLength(usize),
    InvalidState(u8),
}

impl Display for VBAVisibility {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Decrypt(e) => write!(f, "{e}"),
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

#[derive(Debug, PartialEq, Eq)]
pub enum VBADecrypt {
    InvalidHex(String),
    TooShort(String),
    Version(u8),
    LengthMismatch(u32, u32),
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
            Self::TooShort(s) => write!(f, "The hex string {s} is too short to be decrypted"),
            Self::Version(v) => {
                write!(
                    f,
                    "VBA Data Encryption Version MUST be 2 according to the spec, not {v}"
                )
            }
            Self::LengthMismatch(length, data) => {
                write!(f, "The length of the decrypted data: {data} does not match decrypted length: {length}")
            }
        }
    }
}
