use nom::{branch::alt, bytes::complete::tag, combinator::value, IResult};

pub type NwLn = u8;

pub fn parse(input: &[u8]) -> IResult<&[u8], NwLn> {
    value(0x0a, alt((tag("\r\n"), tag("\n\r"))))(input)
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
        assert_eq!(parse(b"\r\n"), Ok((&b""[..], b'\n')));
        assert_eq!(parse(b"\n\r"), Ok((&b""[..], b'\n')));
    }

    #[test]
    fn further_data() {
        assert_eq!(
            parse(b"\r\nsomething else"),
            Ok((&b"something else"[..], b'\n'))
        );
    }

    #[test]
    fn just_one() {
        assert_eq!(
            parse(b"\n"),
            Err(Err::Error(Error::new(&b"\n"[..], ErrorKind::Tag)))
        );
        assert_eq!(
            parse(b"\nsomething else"),
            Err(Err::Error(Error::new(
                &b"\nsomething else"[..],
                ErrorKind::Tag
            )))
        );
    }

    #[test]
    fn other_data_first() {
        assert_eq!(
            parse(b"test\r\nsomething else"),
            Err(Err::Error(Error::new(
                &b"test\r\nsomething else"[..],
                ErrorKind::Tag
            )))
        );
    }
}
