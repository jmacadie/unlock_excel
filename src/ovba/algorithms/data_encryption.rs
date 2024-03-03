//! VBA reversible encryption algorithm
//!
//! Specification can be found [here](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/a02dfe4e-3c9f-45a4-8f14-f2f2d44fa063)
use crate::error;

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
pub fn decode<D: AsRef<[u8]>>(encrypted_data: D) -> Result<Vec<u8>, error::DataEncryption> {
    let encrypted_data = encrypted_data.as_ref();
    if encrypted_data.len() < 8 {
        // 3 for seed, version & project key + 0 ignored + 4 length + 1 data
        let string = encrypted_data
            .iter()
            .fold(String::new(), |s, b| format!("{s}{b:02x}"));
        return Err(error::DataEncryption::TooShort(string));
    }

    let seed = encrypted_data[0];
    let version_enc = encrypted_data[1];
    let project_key_enc = encrypted_data[2];

    let version = seed ^ version_enc;
    if version != 2 {
        return Err(error::DataEncryption::Version(version));
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
        .map_err(|_| error::DataEncryption::LengthMismatch(u32::MAX, length))?;
    if data_len != length {
        return Err(error::DataEncryption::LengthMismatch(data_len, length));
    }

    Ok(data)
}

#[allow(dead_code)]
/// Apply VBA encryption algorithm to a slice of bytes of data
///
/// # Reference
/// Specification can be found [here](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/1ad481e0-7df4-4cac-a9a4-9c29a1340123)
pub fn encode<D: AsRef<[u8]>>(seed: u8, project_key: u8, data: D) -> Vec<u8> {
    const VERSION: u8 = 2;
    let data = data.as_ref();

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

    encrypted_data
}

#[cfg(test)]
mod tests {
    use crate::error;

    use super::*;

    #[test]
    fn decrypt_too_short() {
        let test = [0x21, 0x34, 0x56, 0x78, 0x4a, 0x3b, 0x2f];
        assert_eq!(
            Err(error::DataEncryption::TooShort(
                test.iter()
                    .fold(String::new(), |s, b| format!("{s}{b:02x}"))
            )),
            decode(test)
        );
        let test = [0x7e, 0x2f];
        assert_eq!(
            Err(error::DataEncryption::TooShort(
                test.iter()
                    .fold(String::new(), |s, b| format!("{s}{b:02x}"))
            )),
            decode(test)
        );
    }

    #[test]
    fn decrypt_version_not_2() {
        let test = [0x01, 0x23, 0x45, 0x67, 0x89, 0xab, 0xcd, 0xef];
        assert_eq!(
            Err(error::DataEncryption::Version(0x01 ^ 0x23)),
            decode(test)
        );
    }

    #[test]
    fn decrypt_data_length_mismatch() {
        let test = [0x11, 0x13, 0xeb, 0x02, 0xfa, 0x02, 0xfa, 0x6d, 0x27];
        assert_eq!(
            Err(error::DataEncryption::LengthMismatch(2, 15)),
            decode(test)
        );
    }

    #[test]
    fn encrypt_and_decrypt() {
        let raw =
            b"When he was nearly thirteen, my brother Jem got his arm badly broken at the elbow.";
        let enc = encode(0x0c, 0x9f, raw);
        let dec = decode(enc).unwrap();
        assert_eq!(&raw[..], &dec);
    }
}
