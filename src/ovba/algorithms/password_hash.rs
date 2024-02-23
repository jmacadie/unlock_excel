//! VBA custom format for storing a password hash
//!
//! Specification can be found [here](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/ec1b8759-522b-46d4-bff5-37ed2b1f2ebb)
use rand::Rng;
use sha1::{Digest, Sha1};

use crate::error;

use super::Data;

pub type Salt = [u8; 4];
pub type Hash = [u8; 20];
// TODO: Should actually take an array of bytes that are MBCS characters encoded using the
// code page specified by PROJECTCODEPAGE
type Password<'a> = &'a str;

/// Retrieve the hash and salt from the VBA format for storing hashed passwords
///
/// This mostly does error checking that the data is well formed (see errors section below). After
/// error checking, there is a simple routine to decode any nulls in the salt and hash parts
///
/// Will error if:
/// - The data slice passed is not 29 bytes
/// - The initial byte, which is reserved, is not 0xff
/// - The final terminator byte is not the null byte
/// - The null decoding of the salt and password hash does not find an encoded 0x01 in the
/// locations that are to be replaced with null
pub fn decode<D: AsRef<[u8]>>(data: D) -> Result<(Salt, Hash), error::PasswordHash> {
    let data = data.as_ref();
    if data.len() != 29 {
        return Err(error::PasswordHash::Length(data.len()));
    }
    if data.first() != Some(0xff).as_ref() {
        return Err(error::PasswordHash::Reserved(data[0]));
    }
    if data.last() != Some(0x00).as_ref() {
        return Err(error::PasswordHash::Terminator(data[28]));
    }

    let mut salt = Salt::default();
    salt.clone_from_slice(&data[4..8]);

    let mut hash = Hash::default();
    hash.clone_from_slice(&data[8..28]);

    // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/5797c2e1-4c86-4f44-89b4-1edb30da00cc
    // Add nulls to salt
    let mut grbitkey = data[1];
    for (i, byte) in salt.iter_mut().enumerate() {
        if grbitkey & 1 == 0 {
            if *byte != 0x01 {
                return Err(error::PasswordHash::SaltNull(salt, i));
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
                return Err(error::PasswordHash::HashNull(hash, i));
            }
            *byte = 0;
        }
        grbithashnull >>= 1;
    }

    Ok((salt, hash))
}

/// Convert references to a salt and hash into an encoded byte stream for storage in the VBA
/// project
///
/// The salt is generic over any type that can produce a reference to a slice of bytes. The hash is
/// constrained to fixed array of 20 bytes type since it is assumed this will always come from the
/// `generate_hash` fuction from this module and that output is guaranteed to be well formed
///
/// The ouput is an owned Vector of bytes, since the inputs are concatenated and potentially
/// modified
///
/// Will error if:
/// - The salt is not 4 bytes long
#[allow(dead_code)]
fn encode<S: AsRef<[u8]>>(salt: S, hash: Hash) -> Result<Data, error::PasswordHashEncode> {
    if salt.as_ref().len() != 4 {
        return Err(error::PasswordHashEncode::SaltLength(salt.as_ref().len()));
    }

    let mut output = Vec::with_capacity(29);
    output.push(0xff);

    let mut grbitkey = 0;
    let mut nulled_salt = Salt::default();
    for (i, b) in nulled_salt.iter_mut().enumerate().rev() {
        grbitkey <<= 1;
        match salt.as_ref().get(i) {
            Some(0x00) => {
                *b = 0x01;
            }
            Some(&v) => {
                grbitkey |= 1;
                *b = v;
            }
            None => {
                unreachable!();
            }
        }
    }

    let mut grbithashnull: u32 = 0;
    let mut nulled_hash = Hash::default();
    for (i, b) in nulled_hash.iter_mut().enumerate().rev() {
        grbithashnull <<= 1;
        match hash.get(i) {
            Some(0x00) => {
                *b = 0x01;
            }
            Some(&v) => {
                grbithashnull |= 1;
                *b = v;
            }
            None => {
                unreachable!();
            }
        }
    }

    let grbit = (((grbithashnull & 0x0f) as u8) << 4) | grbitkey;
    output.push(grbit);
    let grbit = ((grbithashnull & 0xff0) >> 4) as u8;
    output.push(grbit);
    let grbit = ((grbithashnull & 0xff000) >> 12) as u8;
    output.push(grbit);

    output.extend_from_slice(&nulled_salt);
    output.extend_from_slice(&nulled_hash);

    output.push(0x00);

    Ok(output.into())
}

/// Generate an SHA1 hash of the given password, expressed in bytes, plus the 4 random bytes of the
/// salt appended to it.
///
/// Outputs a fixed 20 byte array
#[allow(dead_code)]
fn generate_hash<S: AsRef<[u8]>>(password: Password, salt: S) -> Hash {
    let mut hasher = Sha1::new();
    let mut salted: Vec<u8> = password.as_bytes().to_owned();
    salted.extend_from_slice(salt.as_ref());
    hasher.update(salted);
    hasher.finalize().into()
}

/// Hashes the supplied password with the salt & then encodes it for storage in the VBA project
///
/// A separate function from `encode_password` to allow encoding from a deterministic salt value
#[allow(dead_code)]
fn encode_password_with_salt<S: AsRef<[u8]>>(
    password: Password,
    salt: S,
) -> Result<Data, error::PasswordHashEncode> {
    let hash = generate_hash(password, &salt);
    encode(salt, hash)
}

/// Hashes the password with a random salt, and then encodes for storing in the VBA file
#[allow(dead_code)]
pub fn encode_password(password: Password) -> Data {
    let mut rng = rand::thread_rng();
    let salt = [rng.gen(), rng.gen(), rng.gen(), rng.gen()];
    encode_password_with_salt(password, salt).expect("the salt is 4 bytes long")
}

/// Determine if a password matches a salt + hash combination
///
/// Separate to `password_match` as it saves the step of decoding, which is useful where this has
/// already taken place or where it is intended to run this multiple times. In the latter case we
/// will want to decode once and the cache the salt and hash
#[allow(dead_code)]
pub fn password_match_hash(test: Password, salt: Salt, hash: Hash) -> bool {
    generate_hash(test, salt) == hash
}

/// Determine if a password matches the encoded password
///
/// Returns an error if the encoded password cannot be decoded, see `decode`
#[allow(dead_code)]
pub fn password_match<D: AsRef<[u8]>>(
    test: Password,
    encoded_password: D,
) -> Result<bool, error::PasswordHash> {
    let (salt, hash) = decode(encoded_password)?;
    Ok(password_match_hash(test, salt, hash))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn encoded_too_short() {
        let data = [0xff, 0x00];
        assert_eq!(Err(error::PasswordHash::Length(2)), decode(data));
        let data = b"This is an array in disguise";
        assert_eq!(Err(error::PasswordHash::Length(28)), decode(data));
    }

    #[test]
    fn encoded_too_long() {
        let data = [0xff; 1_000];
        assert_eq!(Err(error::PasswordHash::Length(1_000)), decode(data));
        let data = b"This is a longer array in disguise";
        assert_eq!(Err(error::PasswordHash::Length(34)), decode(data));
    }

    #[test]
    fn bad_start_value() {
        let reserved = [0xfe];
        let grbits = [0b1111_1111, 0b1111_1111, 0b1111_1111];
        let salt = [0x12, 0x34, 0x56, 0x78];
        let hash = [
            0x11, 0x22, 0x33, 0x44, 0x55, 0x66, 0x77, 0x88, 0x99, 0xaa, 0xbb, 0xcc, 0xdd, 0xee,
            0xff, 0x11, 0x22, 0x33, 0x44, 0x55,
        ];
        let terminator = [0x00];

        let mut data = Vec::new();
        data.extend_from_slice(&reserved);
        data.extend_from_slice(&grbits);
        data.extend_from_slice(&salt);
        data.extend_from_slice(&hash);
        data.extend_from_slice(&terminator);

        assert_eq!(Err(error::PasswordHash::Reserved(0xfe)), decode(&data));
    }

    #[test]
    fn bad_end_value() {
        let reserved = [0xff];
        let grbits = [0b1111_1111, 0b1111_1111, 0b1111_1111];
        let salt = [0x12, 0x34, 0x56, 0x78];
        let hash = [
            0x11, 0x22, 0x33, 0x44, 0x55, 0x66, 0x77, 0x88, 0x99, 0xaa, 0xbb, 0xcc, 0xdd, 0xee,
            0xff, 0x11, 0x22, 0x33, 0x44, 0x55,
        ];
        let terminator = [0x01];

        let mut data = Vec::new();
        data.extend_from_slice(&reserved);
        data.extend_from_slice(&grbits);
        data.extend_from_slice(&salt);
        data.extend_from_slice(&hash);
        data.extend_from_slice(&terminator);

        assert_eq!(Err(error::PasswordHash::Terminator(0x01)), decode(&data));
    }

    #[test]
    fn bad_salt_null() {
        let reserved = [0xff];
        let grbits = [0b1111_1101, 0b1111_1111, 0b1111_1111];
        let salt = [0x12, 0x34, 0x56, 0x78];
        let hash = [
            0x11, 0x22, 0x33, 0x44, 0x55, 0x66, 0x77, 0x88, 0x99, 0xaa, 0xbb, 0xcc, 0xdd, 0xee,
            0xff, 0x11, 0x22, 0x33, 0x44, 0x55,
        ];
        let terminator = [0x00];

        let mut data = Vec::new();
        data.extend_from_slice(&reserved);
        data.extend_from_slice(&grbits);
        data.extend_from_slice(&salt);
        data.extend_from_slice(&hash);
        data.extend_from_slice(&terminator);

        assert_eq!(Err(error::PasswordHash::SaltNull(salt, 1)), decode(&data));
    }

    #[test]
    fn bad_hash_null() {
        let reserved = [0xff];
        let grbits = [0b1011_1111, 0b1111_1111, 0b1111_1111];
        let salt = [0x12, 0x34, 0x56, 0x78];
        let hash = [
            0x11, 0x22, 0x33, 0x44, 0x55, 0x66, 0x77, 0x88, 0x99, 0xaa, 0xbb, 0xcc, 0xdd, 0xee,
            0xff, 0x11, 0x22, 0x33, 0x44, 0x55,
        ];
        let terminator = [0x00];

        let mut data = Vec::new();
        data.extend_from_slice(&reserved);
        data.extend_from_slice(&grbits);
        data.extend_from_slice(&salt);
        data.extend_from_slice(&hash);
        data.extend_from_slice(&terminator);
        assert_eq!(Err(error::PasswordHash::HashNull(hash, 2)), decode(&data));

        let grbits = [0b1111_1111, 0b1111_1011, 0b1111_1111];
        data.clear();
        data.extend_from_slice(&reserved);
        data.extend_from_slice(&grbits);
        data.extend_from_slice(&salt);
        data.extend_from_slice(&hash);
        data.extend_from_slice(&terminator);
        assert_eq!(Err(error::PasswordHash::HashNull(hash, 6)), decode(&data));

        let grbits = [0b1111_1111, 0b1111_1111, 0b0111_1111];
        data.clear();
        data.extend_from_slice(&reserved);
        data.extend_from_slice(&grbits);
        data.extend_from_slice(&salt);
        data.extend_from_slice(&hash);
        data.extend_from_slice(&terminator);
        assert_eq!(Err(error::PasswordHash::HashNull(hash, 19)), decode(&data));
    }

    #[test]
    fn ok_no_null() {
        let reserved = [0xff];
        let grbits = [0b1111_1111, 0b1111_1111, 0b1111_1111];
        let salt = [0x12, 0x34, 0x56, 0x78];
        let hash = [
            0x11, 0x22, 0x33, 0x44, 0x55, 0x66, 0x77, 0x88, 0x99, 0xaa, 0xbb, 0xcc, 0xdd, 0xee,
            0xff, 0x11, 0x22, 0x33, 0x44, 0x55,
        ];
        let terminator = [0x00];

        let mut data = Vec::new();
        data.extend_from_slice(&reserved);
        data.extend_from_slice(&grbits);
        data.extend_from_slice(&salt);
        data.extend_from_slice(&hash);
        data.extend_from_slice(&terminator);

        let (ds, dh) = decode(&data).unwrap();
        assert_eq!(salt, ds);
        assert_eq!(hash, dh);
    }

    #[test]
    fn ok_with_null() {
        let reserved = [0xff];
        let grbits = [0b1111_1011, 0b0111_1111, 0b1111_1111];
        let mut salt = [0x12, 0x34, 0x01, 0x78];
        let mut hash = [
            0x11, 0x22, 0x33, 0x44, 0x55, 0x66, 0x77, 0x88, 0x99, 0xaa, 0xbb, 0x01, 0xdd, 0xee,
            0xff, 0x11, 0x22, 0x33, 0x44, 0x55,
        ];
        let terminator = [0x00];

        let mut data = Vec::new();
        data.extend_from_slice(&reserved);
        data.extend_from_slice(&grbits);
        data.extend_from_slice(&salt);
        data.extend_from_slice(&hash);
        data.extend_from_slice(&terminator);

        let (ds, dh) = decode(&data).unwrap();
        salt[2] = 0x00;
        hash[11] = 0x00;
        assert_eq!(salt, ds);
        assert_eq!(hash, dh);
    }

    #[test]
    fn encode_and_decode() {
        let salt = [0x12, 0x34, 0x56, 0x78];
        let hash = [
            0x11, 0x22, 0x33, 0x44, 0x55, 0x66, 0x77, 0x88, 0x99, 0xaa, 0xbb, 0xcc, 0xdd, 0xee,
            0xff, 0x11, 0x22, 0x33, 0x44, 0x55,
        ];

        let enc = encode(salt, hash).unwrap();
        let (ds, dh) = decode(enc).unwrap();

        assert_eq!(salt, ds);
        assert_eq!(hash, dh);
    }

    #[test]
    fn encode_and_decode_with_nulls() {
        let salt = [0x12, 0x00, 0x56, 0x00];
        let hash = [
            0x00, 0x22, 0x00, 0x44, 0x55, 0x66, 0x77, 0x88, 0x99, 0xaa, 0xbb, 0xcc, 0xdd, 0x00,
            0xff, 0x11, 0x22, 0x33, 0x00, 0x55,
        ];

        let enc = encode(salt, hash).unwrap();
        let (ds, dh) = decode(enc).unwrap();

        assert_eq!(salt, ds);
        assert_eq!(hash, dh);
    }

    #[test]
    fn encoded_password_matches() {
        let password = "CorrectHorseBatteryStaple";
        let enc = encode_password(password);
        assert!(password_match(password, enc).unwrap());
        let password = "P@ssw0rd";
        let enc = encode_password(password);
        assert!(password_match(password, enc).unwrap());
    }

    #[test]
    fn encoded_password_matches_no_random() {
        let salt = [0x4a, 0x4d, 0x2a, 0x15];
        let password = "CorrectHorseBatteryStaple";
        let enc = encode_password_with_salt(password, salt).unwrap();
        assert!(password_match(password, enc).unwrap());
        let password = "P@ssw0rd";
        let enc = encode_password_with_salt(password, salt).unwrap();
        assert!(password_match(password, enc).unwrap());
    }
}
