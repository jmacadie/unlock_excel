//! VBA reversible encryption algorithm
//!
//! Specification can be found [here](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/a02dfe4e-3c9f-45a4-8f14-f2f2d44fa063)
use std::{fmt::Write, ops::Deref, str::FromStr};

use crate::error;

#[derive(Debug, PartialEq, Eq)]
pub struct Data(Vec<u8>);

impl FromStr for Data {
    type Err = error::VBADecrypt;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        if s.bytes().any(|x| !x.is_ascii_hexdigit()) {
            return Err(error::VBADecrypt::InvalidHex(s.to_owned()));
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

// TODO: Remove? This is not a smart pointer
impl Deref for Data {
    type Target = Vec<u8>;

    fn deref(&self) -> &Self::Target {
        &self.0
    }
}

/// Apply VBA decryption algorithm to a hexadecimal string of encrypted data
///
/// # Reference
/// Specification can be found [here](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/7e9d84fe-86e3-46d6-aaff-8388e72c0168)
///
/// # Error
/// Will generate an error if:
/// - any of the characters are not ascii hex digits: 0-9, a-z, A-Z
/// - the input is too short to correctly contain encrypted data.
/// At the very least, 6 chars are needed for the seed, version &
/// project key, plus 0 to 6 characters of ignored data, plus 8
/// characters for the data length and then at least 2 characters
/// for the data itself. This makes a minimum of 16 characters
/// - the version is not 2. According to the spec the version
/// MUST be 2
/// - the length of the decrypted data does not match the decrypted
/// length parameter
pub fn decrypt_str(hex: &str) -> Result<Data, error::VBADecrypt> {
    let data: Data = hex.parse()?;
    decrypt(&data.0)
}

/// Apply VBA decryption algorithm to a slice of bytes of encrypted data
///
/// # Reference
/// Specification can be found [here](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/7e9d84fe-86e3-46d6-aaff-8388e72c0168)
///
/// # Error
/// Will generate an error if:
/// - the input is too short to correctly contain encrypted data.
/// At the very least, 3 bytes are needed for the seed, version &
/// project key, plus 0 to 3 bytes of ignored data, plus 4
/// bytes for the data length and then at least 2 bytes
/// for the data itself. This makes a minimum of 8 bytes
/// - the version is not 2. According to the spec the version
/// MUST be 2
/// - the length of the decrypted data does not match the decrypted
/// length parameter
pub fn decrypt(encrypted_data: &[u8]) -> Result<Data, error::VBADecrypt> {
    if encrypted_data.len() < 8 {
        // 3 for seed, version & project key + 0 ignored + 4 length + 1 data
        let string = encrypted_data.iter().fold(String::new(), |mut output, b| {
            let _ = write!(output, "{b:02x}");
            output
        });
        return Err(error::VBADecrypt::TooShort(string));
    }

    let seed = encrypted_data[0];
    let version_enc = encrypted_data[1];
    let project_key_enc = encrypted_data[2];

    let version = seed ^ version_enc;
    if version != 2 {
        return Err(error::VBADecrypt::Version(version));
    }
    let project_key = seed ^ project_key_enc;
    let ignored_length = ((seed & 6) >> 1).into();

    let mut unencrypted_byte_1 = project_key;
    let mut encrypted_byte_1 = project_key_enc;
    let mut encrypted_byte_2 = version_enc;

    // Generate the length & data
    let mut data = Vec::new();
    let mut length = 0;
    for (i, byte_enc) in encrypted_data[3..].iter().enumerate() {
        let byte = byte_enc ^ (encrypted_byte_2.wrapping_add(unencrypted_byte_1));
        encrypted_byte_2 = encrypted_byte_1;
        encrypted_byte_1 = *byte_enc;
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

    let data_len = u32::try_from(data.len())
        .map_err(|_| error::VBADecrypt::LengthMismatch(u32::MAX, length))?;
    if data_len != length {
        return Err(error::VBADecrypt::LengthMismatch(data_len, length));
    }

    Ok(Data(data))
}

#[allow(dead_code)]
/// Apply VBA encryption algorithm to a slice of bytes of data
///
/// # Reference
/// Specification can be found [here](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/1ad481e0-7df4-4cac-a9a4-9c29a1340123)
pub fn encrypt(seed: u8, project_key: u8, data: &[u8]) -> Data {
    const VERSION: u8 = 2;
    let version_enc = seed ^ VERSION;
    let project_key_enc = seed ^ project_key;
    let ignored_length = (seed & 6) >> 1;

    let encrypted_length = 3 + usize::from(ignored_length) + 4 + data.len();
    let mut encrypted_data = Vec::with_capacity(encrypted_length);
    encrypted_data.push(seed);
    encrypted_data.push(version_enc);
    encrypted_data.push(project_key_enc);

    let mut unencrypted_byte_1 = project_key;
    let mut encrypted_byte_1 = project_key_enc;
    let mut encrypted_byte_2 = version_enc;

    // Possible trucation Ok as the spec for the algorithm only has a 4 byte integer for length
    #[allow(clippy::cast_possible_truncation)]
    let length = data.len() as u32;

    for byte in (0..ignored_length)
        // spec says any (assume random), but want to be deterministic
        .map(|i| (i * 0x0f) ^ 0xa9)
        .chain(length.to_le_bytes())
        .chain(data.iter().copied())
    {
        let byte_enc = byte ^ (encrypted_byte_2.wrapping_add(unencrypted_byte_1));
        encrypted_data.push(byte_enc);
        encrypted_byte_2 = encrypted_byte_1;
        encrypted_byte_1 = byte_enc;
        unencrypted_byte_1 = byte;
    }

    Data(encrypted_data)
}

#[cfg(test)]
mod tests {
    use crate::error;

    use super::*;

    #[test]
    fn decrypt_non_hex() {
        let test = "123456789abcdefg";
        assert_eq!(
            Err(error::VBADecrypt::InvalidHex(String::from(test))),
            decrypt_str(test)
        );
        let test = "0e2!";
        assert_eq!(
            Err(error::VBADecrypt::InvalidHex(String::from(test))),
            decrypt_str(test)
        );
    }

    #[test]
    fn decrypt_too_short() {
        let test = "123456789abcde";
        assert_eq!(
            Err(error::VBADecrypt::TooShort(String::from(test))),
            decrypt_str(test)
        );
        let test = "0e2f";
        assert_eq!(
            Err(error::VBADecrypt::TooShort(String::from(test))),
            decrypt_str(test)
        );
    }

    #[test]
    fn decrypt_version_not_2() {
        let test = "0123456789abcdef";
        assert_eq!(
            Err(error::VBADecrypt::Version(0x01 ^ 0x23)),
            decrypt_str(test)
        );
    }

    #[test]
    fn decrypt_data_length_mismatch() {
        let test = "1113eb02fa02fa6d27";
        assert_eq!(
            Err(error::VBADecrypt::LengthMismatch(2, 15)),
            decrypt_str(test)
        );
    }

    #[test]
    fn encrypt_and_decrypt() {
        let raw =
            b"When he was nearly thirteen, my brother Jem got his arm badly broken at the elbow.";
        let enc = encrypt(0x0c, 0x9f, raw);
        let dec = decrypt(&enc).unwrap();
        assert_eq!(Vec::from(raw), dec.0);
    }

    #[test]
    fn upper_and_lowercase_hex() {
        let raw = b"It was a bright cold day in April, and the clocks were striking thirteen.";
        let enc = encrypt(0x99, 0xa1, raw);

        let upper = enc.0.iter().fold(String::new(), |mut s, b| {
            let _ = write!(s, "{b:02X}");
            s
        });
        let dec = decrypt_str(&upper).unwrap();
        assert_eq!(Vec::from(raw), dec.0);

        let lower = enc.0.iter().fold(String::new(), |mut s, b| {
            let _ = write!(s, "{b:02x}");
            s
        });
        let dec = decrypt_str(&lower).unwrap();
        assert_eq!(Vec::from(raw), dec.0);
    }
}
