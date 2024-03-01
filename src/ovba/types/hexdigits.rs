use nom::{bytes::complete::take_while_m_n, combinator::map, multi::many_m_n, IResult};

pub type HexDigits = Vec<u8>;

pub fn parse(min: usize, max: usize) -> impl Fn(&[u8]) -> IResult<&[u8], HexDigits> {
    move |input: &[u8]| many_m_n(min / 2, max / 2, parse_hex_pair)(input)
}

fn parse_hex_pair(input: &[u8]) -> IResult<&[u8], u8> {
    map(
        take_while_m_n(2, 2, |b: u8| b.is_ascii_hexdigit()),
        hex_from_u8_slice,
    )(input)
}

// WARN: assumes this will only ever be called by `parse_hex_pair` above
// In particular we're assuming we always get a 2 element slice that only contains ASCII hexdigits.
// If this is not guaranteed, the function really ought to return a Result
fn hex_from_u8_slice(input: &[u8]) -> u8 {
    let upper = match input[0] {
        d if d.is_ascii_digit() => d - b'0',
        c if (b'a'..=b'f').contains(&c) => c - b'a' + 10,
        c if (b'A'..=b'F').contains(&c) => c - b'A' + 10,
        _ => unreachable!(),
    };
    let lower = match input[1] {
        d if d.is_ascii_digit() => d - b'0',
        c if (b'a'..=b'f').contains(&c) => c - b'a' + 10,
        c if (b'A'..=b'F').contains(&c) => c - b'A' + 10,
        _ => unreachable!(),
    };
    (upper << 4) | lower
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
            parse(0, 20)(&b"a1dc9183"[..]),
            Ok((&b""[..], vec![0xa1, 0xdc, 0x91, 0x83]))
        );
        assert_eq!(
            parse(0, 20)(&b"A1DC9183"[..]),
            Ok((&b""[..], vec![0xa1, 0xdc, 0x91, 0x83]))
        );
    }

    #[test]
    fn empty() {
        assert_eq!(parse(0, 20)(&b""[..]), Ok((&b""[..], vec![])));
    }

    #[test]
    fn odd_number_input() {
        assert_eq!(
            parse(0, 20)(&b"a1dc9183c"[..]),
            Ok((&b"c"[..], vec![0xa1, 0xdc, 0x91, 0x83]))
        );
        assert_eq!(parse(0, 20)(&b"3"[..]), Ok((&b"3"[..], vec![])));
    }

    #[test]
    fn too_short() {
        assert_eq!(
            parse(16, 20)(&b"a1dc9183"[..]),
            Err(Err::Error(Error::new(&b""[..], ErrorKind::TakeWhileMN)))
        );
    }

    #[test]
    fn further_data() {
        assert_eq!(
            parse(0, 6)(&b"a1dc9183"[..]),
            Ok((&b"83"[..], vec![0xa1, 0xdc, 0x91]))
        );
        assert_eq!(
            parse(0, 20)(&b"a1dc9183\r\n012345"[..]),
            Ok((&b"\r\n012345"[..], vec![0xa1, 0xdc, 0x91, 0x83]))
        );
    }
}
