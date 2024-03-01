use nom::{
    bytes::complete::{tag, take},
    IResult,
};

use super::i32_from_hex_bytes;

pub type HexInt32 = i32;

pub fn parse(input: &[u8]) -> IResult<&[u8], HexInt32> {
    let (input, _) = tag(b"&H")(input)?;
    let (input, num) = take(8_usize)(input)?;
    let num = i32_from_hex_bytes(num)?;
    Ok((input, num))
}

#[cfg(test)]
mod tests {
    use super::*;
    use nom::{
        error::{Error, ErrorKind},
        Err,
    };

    #[test]
    fn well_formed() {
        assert_eq!(parse(b"&H00000000"), Ok((&b""[..], 0)));
        assert_eq!(
            parse(b"&H7A12CF0A"),
            Ok((&b""[..], i32::from_str_radix("7a12cf0a", 16).unwrap()))
        );
    }

    #[test]
    fn further_data() {
        assert_eq!(
            parse(b"&H7A12CF0A0122"),
            Ok((&b"0122"[..], i32::from_str_radix("7a12cf0a", 16).unwrap()))
        );
    }

    #[test]
    fn missing_opening_tag() {
        assert_eq!(
            parse(b"H7A12CF0A"),
            Err(Err::Error(Error::new(&b"H7A12CF0A"[..], ErrorKind::Tag)))
        );
        assert_eq!(
            parse(b"&7A12CF0A"),
            Err(Err::Error(Error::new(&b"&7A12CF0A"[..], ErrorKind::Tag)))
        );
        assert_eq!(
            parse(b"7A12CF0A"),
            Err(Err::Error(Error::new(&b"7A12CF0A"[..], ErrorKind::Tag)))
        );
    }

    #[test]
    fn too_short() {
        assert_eq!(
            parse(b"&H7A12CF"),
            Err(Err::Error(Error::new(&b"7A12CF"[..], ErrorKind::Eof)))
        );
        assert_eq!(
            parse(b"&H"),
            Err(Err::Error(Error::new(&b""[..], ErrorKind::Eof)))
        );
    }

    #[test]
    fn not_hex() {
        assert_eq!(
            parse(b"&H7A12CF!A"),
            Err(Err::Error(Error::new(
                &b"7A12CF!A"[..],
                ErrorKind::HexDigit
            )))
        );
        assert_eq!(
            parse(b"&H7A 2CF0A"),
            Err(Err::Error(Error::new(
                &b"7A 2CF0A"[..],
                ErrorKind::HexDigit
            )))
        );
    }
}
