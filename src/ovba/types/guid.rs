use nom::{
    bytes::complete::{tag, take},
    IResult,
};

use super::u128_from_hex_bytes;

pub type Guid = u128;

pub fn parse(input: &[u8]) -> IResult<&[u8], Guid> {
    let (input, _) = tag(b"{")(input)?;

    let (input, num) = take(8_usize)(input)?;
    let num = u128_from_hex_bytes(num)?;
    let mut output = num;

    let (input, _) = tag(b"-")(input)?;
    let (input, num) = take(4_usize)(input)?;
    let num = u128_from_hex_bytes(num)?;
    output <<= 16;
    output |= num;

    let (input, _) = tag(b"-")(input)?;
    let (input, num) = take(4_usize)(input)?;
    let num = u128_from_hex_bytes(num)?;
    output <<= 16;
    output |= num;

    let (input, _) = tag(b"-")(input)?;
    let (input, num) = take(4_usize)(input)?;
    let num = u128_from_hex_bytes(num)?;
    output <<= 16;
    output |= num;

    let (input, _) = tag(b"-")(input)?;
    let (input, num) = take(12_usize)(input)?;
    let num = u128_from_hex_bytes(num)?;
    output <<= 48;
    output |= num;

    let (input, _) = tag(b"}")(input)?;

    Ok((input, output))
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
        assert_eq!(
            parse(b"{00000000-0000-0000-0000-000000000000}"),
            Ok((&b""[..], 0))
        );
        assert_eq!(
            parse(b"{3832D640-CF90-11CF-8E43-00A0C911005A}"),
            Ok((
                &b""[..],
                u128::from_str_radix("3832d640cf9011cf8e4300a0c911005a", 16).unwrap()
            ))
        );
    }

    #[test]
    fn further_data() {
        assert_eq!(
            parse(b"{3832D640-CF90-11CF-8E43-00A0C911005A}{00000000-0000-0000-0000-000000000000}"),
            Ok((
                &b"{00000000-0000-0000-0000-000000000000}"[..],
                u128::from_str_radix("3832d640cf9011cf8e4300a0c911005a", 16).unwrap()
            ))
        );
    }

    #[test]
    fn missing_start() {
        assert_eq!(
            parse(b"3832D640-CF90-11CF-8E43-00A0C911005A}"),
            Err(Err::Error(Error::new(
                &b"3832D640-CF90-11CF-8E43-00A0C911005A}"[..],
                ErrorKind::Tag
            )))
        );
    }

    #[test]
    fn missing_end() {
        assert_eq!(
            parse(b"{3832D640-CF90-11CF-8E43-00A0C911005A"),
            Err(Err::Error(Error::new(&b""[..], ErrorKind::Tag)))
        );
    }

    #[test]
    fn missing_hyphen() {
        assert_eq!(
            parse(b"{3832D640CF90-11CF-8E43-00A0C911005A}"),
            Err(Err::Error(Error::new(
                &b"CF90-11CF-8E43-00A0C911005A}"[..],
                ErrorKind::Tag
            )))
        );
        assert_eq!(
            parse(b"{3832D640-CF9011CF-8E43-00A0C911005A}"),
            Err(Err::Error(Error::new(
                &b"11CF-8E43-00A0C911005A}"[..],
                ErrorKind::Tag
            )))
        );
        assert_eq!(
            parse(b"{3832D640-CF90-11CF8E43-00A0C911005A}"),
            Err(Err::Error(Error::new(
                &b"8E43-00A0C911005A}"[..],
                ErrorKind::Tag
            )))
        );
        assert_eq!(
            parse(b"{3832D640-CF90-11CF-8E4300A0C911005A}"),
            Err(Err::Error(Error::new(
                &b"00A0C911005A}"[..],
                ErrorKind::Tag
            )))
        );
    }

    #[test]
    fn missing_numbers() {
        assert_eq!(
            parse(b"{3832D64-CF90-11CF-8E43-00A0C911005A}"),
            Err(Err::Error(Error::new(
                &b"3832D64-"[..],
                ErrorKind::HexDigit
            )))
        );
        assert_eq!(
            parse(b"{3832D640-CF9-11CF-8E43-00A0C911005A}"),
            Err(Err::Error(Error::new(&b"CF9-"[..], ErrorKind::HexDigit)))
        );
        assert_eq!(
            parse(b"{3832D640-CF90-11C-8E43-00A0C911005A}"),
            Err(Err::Error(Error::new(&b"11C-"[..], ErrorKind::HexDigit)))
        );
        assert_eq!(
            parse(b"{3832D640-CF90-11CF-8E4-00A0C911005A}"),
            Err(Err::Error(Error::new(&b"8E4-"[..], ErrorKind::HexDigit)))
        );
        assert_eq!(
            parse(b"{3832D640-CF90-11CF-8E43-00A0C911005}"),
            Err(Err::Error(Error::new(
                &b"00A0C911005}"[..],
                ErrorKind::HexDigit
            )))
        );
    }

    #[test]
    fn extra_numbers() {
        assert_eq!(
            parse(b"{3832D6402-CF90-11CF-8E43-00A0C911005A}"),
            Err(Err::Error(Error::new(
                &b"2-CF90-11CF-8E43-00A0C911005A}"[..],
                ErrorKind::Tag
            )))
        );
        assert_eq!(
            parse(b"{3832D640-CF903-11CF-8E43-00A0C911005A}"),
            Err(Err::Error(Error::new(
                &b"3-11CF-8E43-00A0C911005A}"[..],
                ErrorKind::Tag
            )))
        );
        assert_eq!(
            parse(b"{3832D640-CF90-11CFA-8E43-00A0C911005A}"),
            Err(Err::Error(Error::new(
                &b"A-8E43-00A0C911005A}"[..],
                ErrorKind::Tag
            )))
        );
        assert_eq!(
            parse(b"{3832D640-CF90-11CF-8E439-00A0C911005A}"),
            Err(Err::Error(Error::new(
                &b"9-00A0C911005A}"[..],
                ErrorKind::Tag
            )))
        );
        assert_eq!(
            parse(b"{3832D640-CF90-11CF-8E43-00A0C911005A0}"),
            Err(Err::Error(Error::new(&b"0}"[..], ErrorKind::Tag)))
        );
    }
}
