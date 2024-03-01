use nom::{branch::alt, bytes::complete::tag, combinator::value, IResult};

pub type Whitespace = u8;

pub fn parse(input: &[u8]) -> IResult<&[u8], Whitespace> {
    alt((value(0x20, tag(&[0x20][..])), value(0x09, tag(&[0x09][..]))))(input)
}

#[cfg(test)]
mod tests {
    use super::*;
    use nom::{
        error::{Error, ErrorKind},
        Err,
    };

    #[test]
    fn find_space() {
        assert_eq!(parse(&b" "[..]), Ok((&b""[..], 0x20)));
    }

    #[test]
    fn find_tab() {
        assert_eq!(parse(&b"\t"[..]), Ok((&b""[..], 0x09)));
    }

    #[test]
    fn find_with_more_data() {
        assert_eq!(
            parse(&b" & then something else"[..]),
            Ok((&b"& then something else"[..], 0x20))
        );
    }

    #[test]
    fn double_wsp() {
        assert_eq!(parse(&b"  "[..]), Ok((&b" "[..], 0x20)));
        assert_eq!(parse(&b" \t"[..]), Ok((&b"\t"[..], 0x20)));
        assert_eq!(parse(&b"\t "[..]), Ok((&b" "[..], 0x09)));
    }

    #[test]
    fn not_wsp() {
        assert_eq!(
            parse(&b"x "[..]),
            Err(Err::Error(Error::new(&b"x "[..], ErrorKind::Tag)))
        );
    }
}
