pub mod guid;
pub mod hex_int_32;
pub mod hexdigits;
pub mod int_32;
pub mod module_identifier;
pub mod new_line;
pub mod path;
pub mod quoted_character;
pub mod quoted_characters;
pub mod whitespace;

// TODO: Turn this into a macro rules expansion
fn i32_from_hex_bytes(bytes: &[u8]) -> Result<i32, nom::Err<nom::error::Error<&[u8]>>> {
    let Ok(num) = std::str::from_utf8(bytes) else {
        return Err(nom::Err::Error(nom::error::Error::new(
            bytes,
            nom::error::ErrorKind::HexDigit,
        )));
    };
    i32::from_str_radix(num, 16).map_err(|_| {
        nom::Err::Error(nom::error::Error::new(
            bytes,
            nom::error::ErrorKind::HexDigit,
        ))
    })
}

fn u128_from_hex_bytes(bytes: &[u8]) -> Result<u128, nom::Err<nom::error::Error<&[u8]>>> {
    let Ok(num) = std::str::from_utf8(bytes) else {
        return Err(nom::Err::Error(nom::error::Error::new(
            bytes,
            nom::error::ErrorKind::HexDigit,
        )));
    };
    u128::from_str_radix(num, 16).map_err(|_| {
        nom::Err::Error(nom::error::Error::new(
            bytes,
            nom::error::ErrorKind::HexDigit,
        ))
    })
}
