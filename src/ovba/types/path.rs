use nom::IResult;

use super::quoted_characters;

pub type Path = String;

pub fn parse(input: &[u8]) -> IResult<&[u8], Path> {
    quoted_characters::parse(0, 259)(input)
}

#[cfg(test)]
mod tests {
    use super::*;
    use nom::{
        error::{Error, ErrorKind},
        Err,
    };

    #[test]
    fn find_a_path() {
        assert_eq!(
            parse(&b"\"C:\\Program Files\\Microsoft Office\\root\\Office16\""[..]),
            Ok((
                &b""[..],
                String::from("C:\\Program Files\\Microsoft Office\\root\\Office16")
            ))
        );
    }

    #[test]
    fn escaped_dquote() {
        assert_eq!(
            parse(&b"\"C:\\Program Files\\Microsoft Office\\\"\"root\"\"\\Office16\""[..]),
            Ok((
                &b""[..],
                String::from("C:\\Program Files\\Microsoft Office\\\"root\"\\Office16")
            ))
        );
    }

    #[test]
    fn missing_start_or_end_dquotes() {
        assert_eq!(
            parse(&b"C:\\Program Files\\Microsoft Office\\root\\Office16\""[..]),
            Err(Err::Error(Error::new(
                &b"C:\\Program Files\\Microsoft Office\\root\\Office16\""[..],
                ErrorKind::Tag
            )))
        );
        assert_eq!(
            parse(&b"\"C:\\Program Files\\Microsoft Office\\root\\Office16"[..]),
            Err(Err::Error(Error::new(&b""[..], ErrorKind::Tag)))
        );
    }

    #[test]
    fn further_data() {
        assert_eq!(
            parse(&b"\"C:\\Program Files\\Microsoft Office\\root\\Office16\" and now for something completely different"[..]),
            Ok((
                &b" and now for something completely different"[..],
                String::from("C:\\Program Files\\Microsoft Office\\root\\Office16")
            ))
        );
    }

    #[test]
    fn invalid_character() {
        assert_eq!(
            parse(&b"\"C:\\Program Files\\Microsoft Office\\ro\not\\Office16\""[..]),
            Err(Err::Error(Error::new(
                &b"\not\\Office16\""[..],
                ErrorKind::Tag
            )))
        );
        assert_eq!(
            parse(&b"\"C:\\Program Files\\Microsoft Office\\ro\0ot\\Office16\""[..]),
            Err(Err::Error(Error::new(
                &b"\0ot\\Office16\""[..],
                ErrorKind::Tag
            )))
        );
    }

    #[test]
    fn too_long() {
        assert_eq!(
            parse(&b"\"C:\\Program Files\\Microsoft Office\\root\\Office16ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff\""[..]),
            Ok(( &b""[..],
                String::from("C:\\Program Files\\Microsoft Office\\root\\Office16ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff"),
            ))
        );
        assert_eq!(
            parse(&b"\"C:\\Program Files\\Microsoft Office\\root\\Office16fffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff\""[..]),
            Err(Err::Error(Error::new(
                &b"f\""[..],
                ErrorKind::Tag
            )))
        );
    }
}
