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
    ProjectStructure(ProjectStructure),
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

impl From<ProtectionState> for UnlockError {
    fn from(value: ProtectionState) -> Self {
        Self::ProjectStructure(ProjectStructure::ProtectionState(value))
    }
}

impl From<Password> for UnlockError {
    fn from(value: Password) -> Self {
        Self::ProjectStructure(ProjectStructure::Password(value))
    }
}

impl From<Visibility> for UnlockError {
    fn from(value: Visibility) -> Self {
        Self::ProjectStructure(ProjectStructure::Visibility(value))
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

pub enum ProjectStructure {
    ProtectionState(ProtectionState),
    Password(Password),
    Visibility(Visibility),
}

impl Display for ProjectStructure {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::ProtectionState(e) => write!(f, "{e}"),
            Self::Password(e) => write!(f, "{e}"),
            Self::Visibility(e) => write!(f, "{e}"),
        }
    }
}

impl From<ProtectionState> for ProjectStructure {
    fn from(value: ProtectionState) -> Self {
        Self::ProtectionState(value)
    }
}

impl From<Password> for ProjectStructure {
    fn from(value: Password) -> Self {
        Self::Password(value)
    }
}

impl From<Visibility> for ProjectStructure {
    fn from(value: Visibility) -> Self {
        Self::Visibility(value)
    }
}

#[derive(Debug, PartialEq, Eq)]
pub enum ProtectionState {
    Decrypt(DataEncryption),
    DataLength(usize),
    ReservedBits([u8; 4]),
}

impl Display for ProtectionState {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Decrypt(e) => write!(f, "{e}"),
            Self::DataLength(l) => write!(f, "The Data array for Data Encryption of the VBA Project Protection State SHOULD be 4 bytes, not {l} bytes"),
            Self::ReservedBits(data) => write!(f, "The upper 29 bits of data are reserved and MUST all be 0. Data decodeded to {:08b}{:08b}{:08b}{:08b}", data[0], data[1], data[2], data[3]),
        }
    }
}

impl From<DataEncryption> for ProtectionState {
    fn from(value: DataEncryption) -> Self {
        Self::Decrypt(value)
    }
}

#[derive(Debug, PartialEq, Eq)]
pub enum Password {
    Decrypt(DataEncryption),
    None(PasswordNone),
    Hash(PasswordHash),
    PlainText(PasswordPlain),
    NoData,
}

impl Display for Password {
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

impl From<DataEncryption> for Password {
    fn from(value: DataEncryption) -> Self {
        Self::Decrypt(value)
    }
}

impl From<PasswordNone> for Password {
    fn from(value: PasswordNone) -> Self {
        Self::None(value)
    }
}

impl From<PasswordHash> for Password {
    fn from(value: PasswordHash) -> Self {
        Self::Hash(value)
    }
}

impl From<PasswordPlain> for Password {
    fn from(value: PasswordPlain) -> Self {
        Self::PlainText(value)
    }
}

#[derive(Debug, PartialEq, Eq)]
pub enum PasswordNone {
    NotNull(u8),
}

impl Display for PasswordNone {
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
pub enum PasswordHash {
    Length(usize),
    Reserved(u8),
    Terminator(u8),
    SaltNull([u8; 4], usize),
    HashNull([u8; 20], usize),
}

impl Display for PasswordHash {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Length(l) => write!(f, "The digest for the password hash must be 29 bytes, not {l}"),
            Self::Reserved(b) => write!(f, "The first byte of the VBA password hash data structure is reserved and MUST be 0xff, not 0x{b:02x}"),
            Self::Terminator(b) => write!(f, "The final byte of the VBA password hash data structure is the terminator and MUST be 0x00, not 0x{b:02x}"),
            Self::SaltNull(data, i) => write!(f, "The byte in position {i} of the salt {data:?} is being replaced with a null. It should have a value of 1 before update"),
            Self::HashNull(data, i) => write!(f, "The byte in position {i} of the hash {data:?} is being replaced with a null. It should have a value of 1 before update"),
        }
    }
}

#[derive(Debug, PartialEq, Eq)]
pub enum PasswordHashEncode {
    SaltLength(usize),
}

impl Display for PasswordHashEncode {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::SaltLength(l) => write!(f, "The salt must be 4 bytes, not {l}"),
        }
    }
}

#[derive(Debug, PartialEq, Eq)]
pub enum PasswordPlain {
    Terminator(u8),
}

impl Display for PasswordPlain {
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
pub enum Visibility {
    Decrypt(DataEncryption),
    DataLength(usize),
    InvalidState(u8),
}

impl Display for Visibility {
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

impl From<DataEncryption> for Visibility {
    fn from(value: DataEncryption) -> Self {
        Self::Decrypt(value)
    }
}

#[derive(Debug, PartialEq, Eq)]
pub enum DataEncryption {
    InvalidHex(InvalidHex),
    TooShort(String),
    Version(u8),
    LengthMismatch(u32, u32),
}

impl Display for DataEncryption {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::InvalidHex(e) => write!(f, "{e}"),
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

impl From<InvalidHex> for DataEncryption {
    fn from(value: InvalidHex) -> Self {
        Self::InvalidHex(value)
    }
}

#[derive(Debug, PartialEq, Eq)]
pub struct InvalidHex(String);

impl From<String> for InvalidHex {
    fn from(value: String) -> Self {
        Self(value)
    }
}

impl From<ParseIntError> for InvalidHex {
    fn from(value: ParseIntError) -> Self {
        Self(format!("{value}"))
    }
}

impl Display for InvalidHex {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            f,
            "Cannot apply VBA data decryption as supplied value is not valid hex: {}",
            self.0
        )
    }
}
