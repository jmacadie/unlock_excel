pub mod data_encryption;
pub mod password_hash;

use crate::error;
use std::fmt::Display;
use std::ops::Deref;
use std::str::FromStr;

/// Common data structure to represent an arbitrary stream of bytes.
///
/// Has been created to easily allow conversion to and from hex string representation of the data,
/// which happens in a few places in this crate
#[derive(Debug, PartialEq, Eq)]
pub struct Data(Vec<u8>);

impl FromStr for Data {
    type Err = error::InvalidHex;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        if s.bytes().any(|x| !x.is_ascii_hexdigit()) {
            return Err(s.to_owned().into());
        }
        let length = s.len() / 2;
        let mut data = Vec::with_capacity(length);
        data.extend(s.as_bytes().chunks_exact(2).map(|x| {
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
        }));
        Ok(Self(data))
    }
}

impl Display for Data {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        for b in &self.0 {
            write!(f, "{b:02X}")?;
        }
        Ok(())
    }
}

impl From<Vec<u8>> for Data {
    fn from(value: Vec<u8>) -> Self {
        Self(value)
    }
}

impl AsRef<[u8]> for Data {
    fn as_ref(&self) -> &[u8] {
        &self.0
    }
}

// TODO: Remove? This is not a smart pointer
impl Deref for Data {
    type Target = Vec<u8>;

    fn deref(&self) -> &Self::Target {
        &self.0
    }
}
