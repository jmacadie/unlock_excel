use crate::ovba::types::whitespace;
use nom::{branch::alt, bytes::complete::take_while_m_n, combinator::value, IResult};

pub type QuotedChar = u8;

pub fn parse(input: &[u8]) -> IResult<&[u8], QuotedChar> {
    alt((parse_nq_char, whitespace::parse, parse_dquote))(input)
}

fn parse_nq_char(input: &[u8]) -> IResult<&[u8], QuotedChar> {
    let (input, nq_char) = take_while_m_n(1, 1, |b: u8| b >= 0x23 || b == 0x21)(input)?;
    Ok((input, nq_char[0]))
}

fn parse_dquote(input: &[u8]) -> IResult<&[u8], QuotedChar> {
    value(0x22, take_while_m_n(2, 2, |b: u8| b == 0x22))(input)
}

#[cfg(test)]
mod tests {
    use super::*;
    use nom::{
        error::{Error, ErrorKind},
        Err,
    };

    #[test]
    fn find_nq_char() {
        assert_eq!(parse(&b"!"[..]), Ok((&b""[..], 0x21)));
        assert_eq!(parse(&b"+"[..]), Ok((&b""[..], 0x2b)));
        assert_eq!(parse(&b"5"[..]), Ok((&b""[..], 0x35)));
        assert_eq!(parse(&b"D"[..]), Ok((&b""[..], 0x44)));
        assert_eq!(parse(&b"m"[..]), Ok((&b""[..], 0x6d)));
        assert_eq!(parse(&[0xac][..]), Ok((&b""[..], 0xac)));
    }

    #[test]
    fn find_wsp() {
        assert_eq!(parse(&b" "[..]), Ok((&b""[..], 0x20)));
        assert_eq!(parse(&b"\t"[..]), Ok((&b""[..], 0x09)));
    }

    #[test]
    fn find_dquote() {
        assert_eq!(parse(&b"\"\""[..]), Ok((&b""[..], 0x22)));
    }

    #[test]
    fn single_dquote() {
        assert_eq!(
            parse(&b"\""[..]),
            Err(Err::Error(Error::new(&b"\""[..], ErrorKind::TakeWhileMN)))
        );
        assert_eq!(
            parse(&b"\" \""[..]),
            Err(Err::Error(Error::new(
                &b"\" \""[..],
                ErrorKind::TakeWhileMN
            )))
        );
    }

    #[test]
    fn further_data() {
        assert_eq!(parse(&b"! "[..]), Ok((&b" "[..], 0x21)));
        assert_eq!(parse(&b"+a"[..]), Ok((&b"a"[..], 0x2b)));
        assert_eq!(parse(&b"55"[..]), Ok((&b"5"[..], 0x35)));
        assert_eq!(parse(&b"D\t"[..]), Ok((&b"\t"[..], 0x44)));
        assert_eq!(parse(&b"m;"[..]), Ok((&b";"[..], 0x6d)));
        assert_eq!(parse(&[0xac, 0x30][..]), Ok((&b"0"[..], 0xac)));
    }

    #[test]
    fn invalid_character() {
        assert_eq!(
            parse(&b"\r"[..]),
            Err(Err::Error(Error::new(&b"\r"[..], ErrorKind::TakeWhileMN)))
        );
        assert_eq!(
            parse(&b"\n"[..]),
            Err(Err::Error(Error::new(&b"\n"[..], ErrorKind::TakeWhileMN)))
        );
        assert_eq!(
            parse(&b"\r\n"[..]),
            Err(Err::Error(Error::new(&b"\r\n"[..], ErrorKind::TakeWhileMN)))
        );
        assert_eq!(
            parse(&b"\0"[..]),
            Err(Err::Error(Error::new(&b"\0"[..], ErrorKind::TakeWhileMN)))
        );
    }
}
