use super::quoted_character;
use nom::{bytes::complete::tag, combinator::map, multi::many_m_n, sequence::delimited, IResult};

pub fn parse(min: usize, max: usize) -> impl Fn(&[u8]) -> IResult<&[u8], String> {
    move |input: &[u8]| {
        map(
            delimited(
                tag("\""),
                many_m_n(min, max, quoted_character::parse),
                tag("\""),
            ),
            // TODO: Meant to support MBCS characters
            // This is pretending we only ever have ASCII here
            // Worse, it will panic if non-ASCII is ever passed
            |p: Vec<u8>| String::from_utf8(p).unwrap(),
        )(input)
    }
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
            parse(1, 20)(&b"\"A quoted string\""[..]),
            Ok((&b""[..], String::from("A quoted string")))
        );
    }

    #[test]
    fn quoted_dquote() {
        assert_eq!(
            parse(1, 20)(&b"\"A \"\"quoted\"\" string\""[..]),
            Ok((&b""[..], String::from("A \"quoted\" string")))
        );
    }

    #[test]
    fn too_short() {
        assert_eq!(
            parse(16, 20)(&b"\"A quoted string\""[..]),
            Err(Err::Error(Error::new(&b"\""[..], ErrorKind::TakeWhileMN)))
        );
    }

    #[test]
    fn too_long() {
        assert_eq!(
            parse(1, 13)(&b"\"A quoted string\""[..]),
            Err(Err::Error(Error::new(&b"ng\""[..], ErrorKind::Tag)))
        );
    }

    #[test]
    fn invalid_character() {
        assert_eq!(
            parse(1, 20)(&b"\"A quo\nted string\""[..]),
            Err(Err::Error(Error::new(
                &b"\nted string\""[..],
                ErrorKind::Tag
            )))
        );
        assert_eq!(
            parse(1, 20)(&b"\"A quoted str\0ing\""[..]),
            Err(Err::Error(Error::new(&b"\0ing\""[..], ErrorKind::Tag)))
        );
    }

    #[test]
    fn further_data() {
        assert_eq!(
            parse(1, 20)(&b"\"A quoted string\" plus a bit more"[..]),
            Ok((&b" plus a bit more"[..], String::from("A quoted string")))
        );
    }
}
