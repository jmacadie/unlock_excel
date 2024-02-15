use crate::consts;
use crate::UnlockResult;
use cfb::Stream;
use std::io::BufRead;

pub fn unlocked_project<T: std::io::Read + std::io::Seek>(
    mut project: Stream<T>,
) -> UnlockResult<Vec<u8>> {
    let mut line = Vec::new();
    let mut output = Vec::new();

    while project.read_until(b'\n', &mut line)? > 0 {
        match line.get(0..5) {
            Some(&[b'I', b'D', b'=', b'"', b'{']) => {
                output.extend_from_slice(consts::UNLOCKED_ID.as_bytes());
            }
            Some(&[b'C', b'M', b'G', b'=', b'"']) => {
                output.extend_from_slice(consts::UNLOCKED_CMG.as_bytes());
            }
            Some(&[b'D', b'P', b'B', b'=', b'"']) => {
                output.extend_from_slice(consts::UNLOCKED_DPB.as_bytes());
            }
            Some(&[b'G', b'C', b'=', b'"', _]) => {
                output.extend_from_slice(consts::UNLOCKED_GC.as_bytes());
            }
            _ => output.extend_from_slice(&line),
        }
        line.clear();
    }

    Ok(output)
}
