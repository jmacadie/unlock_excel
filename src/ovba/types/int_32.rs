use nom::{
    bytes::complete::tag,
    character::complete::digit1,
    combinator::{map_res, opt, recognize},
    sequence::pair,
    IResult,
};

pub type Int32 = i32;

pub fn parse(input: &[u8]) -> IResult<&[u8], Int32> {
    map_res(recognize(pair(opt(tag("-")), digit1)), |n: &[u8]| {
        std::str::from_utf8(n)
            .expect("ASCII numbers and - are convertible to UTF-8")
            .parse()
    })(input)
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
        assert_eq!(parse(b"0"), Ok((&b""[..], 0)));
        assert_eq!(parse(b"12340"), Ok((&b""[..], 12340)));
        assert_eq!(parse(b"-7789"), Ok((&b""[..], -7789)));
    }

    #[test]
    fn further_data() {
        assert_eq!(parse(b"12340\r\n"), Ok((&b"\r\n"[..], 12340)));
    }

    #[test]
    fn too_large() {
        assert_eq!(parse(b"2147483647"), Ok((&b""[..], 2_147_483_647)));
        assert_eq!(parse(b"-2147483648"), Ok((&b""[..], -2_147_483_648)));
        assert_eq!(
            parse(b"2147483648"),
            Err(Err::Error(Error::new(
                &b"2147483648"[..],
                ErrorKind::MapRes
            )))
        );
        assert_eq!(
            parse(b"-2147483649"),
            Err(Err::Error(Error::new(
                &b"-2147483649"[..],
                ErrorKind::MapRes
            )))
        );
    }

    #[test]
    fn invalid_character() {
        assert_eq!(
            parse(b".24"),
            Err(Err::Error(Error::new(&b".24"[..], ErrorKind::Digit)))
        );
    }
}
