use nom::{
    bytes::complete::take_while_m_n,
    character::{is_alphabetic, is_alphanumeric},
    combinator::{map, recognize},
    sequence::pair,
    IResult,
};

pub type ModuleIdentifier = String;

pub fn parse(input: &[u8]) -> IResult<&[u8], ModuleIdentifier> {
    map(
        recognize(pair(
            take_while_m_n(1, 1, is_alphabetic),
            take_while_m_n(0, 30, |b| is_alphanumeric(b) || b == b'_'),
        )),
        |s: &[u8]| {
            String::from_utf8(s.to_vec())
                .expect("alphanumeric bytes and _ converting into a String")
        },
    )(input)
}

#[cfg(test)]
mod tests {
    use super::*;
    use nom::{
        error::{Error, ErrorKind},
        Err,
    };

    #[test]
    fn bad_leading_char() {
        assert_eq!(
            parse(b"01234"),
            Err(Err::Error(Error::new(
                &b"01234"[..],
                ErrorKind::TakeWhileMN
            )))
        );
        assert_eq!(
            parse(b" 1234"),
            Err(Err::Error(Error::new(
                &b" 1234"[..],
                ErrorKind::TakeWhileMN
            )))
        );
        assert_eq!(
            parse(b"*1234"),
            Err(Err::Error(Error::new(
                &b"*1234"[..],
                ErrorKind::TakeWhileMN
            )))
        );
    }

    #[test]
    fn long_input() {
        assert_eq!(
            parse(b"A_really_really_long_string_that_is_more_than_31_characters"),
            Ok((
                &b"t_is_more_than_31_characters"[..],
                String::from("A_really_really_long_string_tha")
            ))
        );
    }

    #[test]
    fn terminating_char() {
        assert_eq!(parse(b"A01234\n"), Ok((&b"\n"[..], String::from("A01234"))));
        assert_eq!(
            parse(b"A01234\r\n"),
            Ok((&b"\r\n"[..], String::from("A01234")))
        );
        assert_eq!(
            parse(b"A01234\n\r"),
            Ok((&b"\n\r"[..], String::from("A01234")))
        );
        assert_eq!(
            parse(b"A01234&another_thing"),
            Ok((&b"&another_thing"[..], String::from("A01234")))
        );
    }

    #[test]
    fn short_input() {
        assert_eq!(parse(b"A_module"), Ok((&b""[..], String::from("A_module"))));
        assert_eq!(
            parse(b"A_really_really_long_module_xxx"),
            Ok((&b""[..], String::from("A_really_really_long_module_xxx")))
        );
    }
}
