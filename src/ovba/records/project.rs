#![allow(clippy::doc_markdown, dead_code)]
//! A Struct to hold the contents of the Project Properties stream
//!
//! The PROJECT stream specifies properties of the VBA project. This stream is an array of bytes that specifies properties of the VBA project.
//!
//! VBAPROJECTText = ProjectId
//!                  *ProjectItem
//!                  [ProjectHelpFile]
//!                  [ProjectExeName32]
//!                  ProjectName
//!                  ProjectHelpId
//!                  [ProjectDescription]
//!                  [ProjectVersionCompat32]
//!                  ProjectProtectionState
//!                  ProjectPassword
//!                  ProjectVisibilityState
//!                  NWLN HostExtenders
//!                  [NWLN ProjectWorkspace]
//!
//! Specification can be found [here](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/d88cb9d8-a475-423d-b370-cc0caaf78628)

use crate::{
    error,
    ovba::types::{guid, hex_int_32, int_32, module_identifier, path},
};
use cfb::Stream;
use nom::Finish;
use std::io::Read;

#[derive(Debug)]
pub struct Project {
    id: guid::Guid,
    items: Vec<Item>,
    help_file: Option<path::Path>,
    exe_name: Option<path::Path>,
    name: ProjectIdentifier,
    help_id: int_32::Int32,
    description: Option<DescriptionText>,
    protection_state: ProtectionState,
    password: Password,
    visibility_state: Visibility,
    host_extenders: Vec<HostExtenderRef>,
    workspace: Option<Vec<WindowRecord>>,
}

// String Types
// Between 1 and 128 characters, must be surrounded by quotes, characters are quotedchar type
type ProjectIdentifier = String;
// Up to 2,000 characters, must be surrounded by quotes, characters are quotedchar type
type DescriptionText = String;
// Any number of bytes above 0x20 (space), that don't include 0x3b (;)
type LibName = String;

#[derive(Debug)]
enum Item {
    Module(Module),
    Package(guid::Guid),
}

#[derive(Debug)]
enum Module {
    Doc(module_identifier::ModuleIdentifier, hex_int_32::HexInt32),
    Std(module_identifier::ModuleIdentifier),
    Class(module_identifier::ModuleIdentifier),
    Designer(module_identifier::ModuleIdentifier),
}

#[derive(Debug)]
struct ProtectionState {
    user: bool,
    host: bool,
    vbe: bool,
}

#[derive(Debug)]
pub enum Password {
    None,
    Hash([u8; 4], [u8; 20]),
    Plain(String),
}

#[derive(Debug)]
enum Visibility {
    NotVisible,
    Visible,
}

#[derive(Debug)]
struct HostExtenderRef {
    index: hex_int_32::HexInt32,
    guid: guid::Guid,
    lib: LibName,
    creation_flags: hex_int_32::HexInt32,
}

#[derive(Debug)]
struct WindowRecord {
    module: module_identifier::ModuleIdentifier,
    code: Window,
    designer: Option<Window>,
}

#[derive(Debug)]
struct Window {
    left: int_32::Int32,
    top: int_32::Int32,
    right: int_32::Int32,
    bottom: int_32::Int32,
    state: WindowState,
}

#[derive(Debug)]
enum WindowState {
    Closed,
    Zoomed,
    Minimized,
}

impl Project {
    pub fn from_stream<T: std::io::Read + std::io::Seek>(
        mut stream: Stream<T>,
    ) -> Result<Self, error::ProjectStructure> {
        let mut buf = Vec::new();
        let _ = stream.read_to_end(&mut buf);

        let (_res, p) = nom_parse::project(&buf)
            .finish()
            .map_err(|e| error::ProjectStructure::NomParseError(e.input.to_vec(), buf.clone()))?;

        Ok(p)
    }

    pub const fn is_locked(&self) -> bool {
        self.protection_state.vbe
    }

    pub const fn password(&self) -> &Password {
        &self.password
    }
}

mod nom_parse {
    use super::{
        DescriptionText, HostExtenderRef, Item, LibName, Module, Password, Project,
        ProjectIdentifier, ProtectionState, Visibility, Window, WindowRecord, WindowState,
    };
    use crate::{
        error,
        ovba::{
            algorithms::{data_encryption, password_hash},
            types::{
                guid, hex_int_32, hexdigits, int_32, module_identifier, new_line, path,
                quoted_characters,
            },
        },
    };
    use nom::{
        branch::alt,
        bytes::complete::{tag, take_while},
        character::complete::one_of,
        combinator::{map, map_res, opt},
        multi::{many0, separated_list0},
        sequence::{delimited, pair, preceded, terminated, tuple},
        IResult,
    };

    pub(super) fn project(input: &[u8]) -> IResult<&[u8], Project> {
        map(
            tuple((
                id,
                items,
                opt(help_file),
                opt(exe_name_32),
                name,
                help_id,
                opt(description),
                opt(version_compat_32),
                protection_state,
                password,
                visibility_state,
                host_extenders,
                opt(workspace),
            )),
            |(
                id,
                items,
                help_file,
                exe_name,
                name,
                help_id,
                description,
                _,
                protection_state,
                password,
                visibility_state,
                host_extenders,
                workspace,
            )| {
                Project {
                    id,
                    items,
                    help_file,
                    exe_name,
                    name,
                    help_id,
                    description,
                    protection_state,
                    password,
                    visibility_state,
                    host_extenders,
                    workspace,
                }
            },
        )(input)
    }

    fn id(input: &[u8]) -> IResult<&[u8], guid::Guid> {
        delimited(tag("ID=\""), guid::parse, pair(tag("\""), new_line::parse))(input)
    }

    fn document_module(input: &[u8]) -> IResult<&[u8], Module> {
        map(
            pair(
                preceded(tag("Document="), module_identifier::parse),
                preceded(tag([0x2f]), hex_int_32::parse),
            ),
            |(module, doc_tlib_ver)| Module::Doc(module, doc_tlib_ver),
        )(input)
    }

    fn std_module(input: &[u8]) -> IResult<&[u8], Module> {
        map(
            preceded(tag("Module="), module_identifier::parse),
            Module::Std,
        )(input)
    }

    fn class_module(input: &[u8]) -> IResult<&[u8], Module> {
        map(
            preceded(tag("Class="), module_identifier::parse),
            Module::Class,
        )(input)
    }

    fn designer_module(input: &[u8]) -> IResult<&[u8], Module> {
        map(
            preceded(tag("BaseClass="), module_identifier::parse),
            Module::Designer,
        )(input)
    }

    fn module(input: &[u8]) -> IResult<&[u8], Item> {
        map(
            alt((document_module, std_module, class_module, designer_module)),
            Item::Module,
        )(input)
    }

    fn package(input: &[u8]) -> IResult<&[u8], Item> {
        map(preceded(tag("Package="), guid::parse), Item::Package)(input)
    }

    fn items(input: &[u8]) -> IResult<&[u8], Vec<Item>> {
        terminated(
            separated_list0(new_line::parse, alt((module, package))),
            new_line::parse,
        )(input)
    }

    fn help_file(input: &[u8]) -> IResult<&[u8], path::Path> {
        delimited(tag("HelpFile="), path::parse, new_line::parse)(input)
    }

    fn exe_name_32(input: &[u8]) -> IResult<&[u8], path::Path> {
        delimited(tag("ExeName32="), path::parse, new_line::parse)(input)
    }

    fn name(input: &[u8]) -> IResult<&[u8], ProjectIdentifier> {
        delimited(
            tag("Name="),
            quoted_characters::parse(1, 128),
            new_line::parse,
        )(input)
    }

    fn help_id(input: &[u8]) -> IResult<&[u8], int_32::Int32> {
        delimited(
            tag("HelpContextID=\""),
            int_32::parse,
            pair(tag("\""), new_line::parse),
        )(input)
    }

    fn description(input: &[u8]) -> IResult<&[u8], DescriptionText> {
        delimited(
            tag("Description="),
            quoted_characters::parse(0, 2000),
            new_line::parse,
        )(input)
    }

    fn version_compat_32(input: &[u8]) -> IResult<&[u8], &[u8]> {
        terminated(tag("VersionCompatible32=\"393222000\""), new_line::parse)(input)
    }

    fn protection_state(input: &[u8]) -> IResult<&[u8], ProtectionState> {
        map_res(
            delimited(
                tag("CMG=\""),
                hexdigits::parse(22, 28),
                pair(tag("\""), new_line::parse),
            ),
            |encrypted: Vec<u8>| {
                let data = data_encryption::decode(encrypted)?.into_inner();
                if data.len() != 4 {
                    return Err(error::ProtectionState::DataLength(data.len()));
                }
                if data[0] > 7 || data[1] != 0 || data[2] != 0 || data[3] != 0 {
                    return Err(error::ProtectionState::ReservedBits([
                        data[0], data[1], data[2], data[3],
                    ]));
                }
                let user_protected = data[0] & 1 == 1;
                let host_protected = data[0] & 2 == 2;
                let vbe_protected = data[0] & 4 == 4;
                Ok(ProtectionState {
                    user: user_protected,
                    host: host_protected,
                    vbe: vbe_protected,
                })
            },
        )(input)
    }

    fn password(input: &[u8]) -> IResult<&[u8], Password> {
        map_res(
            delimited(
                tag("DPB=\""),
                hexdigits::parse(16, 2000),
                pair(tag("\""), new_line::parse),
            ),
            |encrypted: Vec<u8>| {
                let data = data_encryption::decode(encrypted)?.into_inner();
                Ok(match data.len() {
                    0 => return Err(error::Password::NoData),
                    1 => {
                        if data.first() != Some(0x00).as_ref() {
                            return Err(error::PasswordNone::NotNull(data[0]).into());
                        }
                        Password::None
                    }
                    29 => {
                        let (salt, hash) = password_hash::decode(data)?;
                        Password::Hash(salt, hash)
                    }
                    _ => {
                        if data.last() != Some(0x00).as_ref() {
                            return Err(error::PasswordPlain::Terminator(*data.last().expect(
                                "Cannot construct a plain password with zero length data",
                            ))
                            .into());
                        }
                        let password =
                            String::from_utf8_lossy(&data[0..(data.len() - 1)]).to_string();
                        Password::Plain(password)
                    }
                })
            },
        )(input)
    }

    fn visibility_state(input: &[u8]) -> IResult<&[u8], Visibility> {
        map_res(
            delimited(
                tag("GC=\""),
                hexdigits::parse(16, 22),
                pair(tag("\""), new_line::parse),
            ),
            |encrypted: Vec<u8>| {
                let data = data_encryption::decode(encrypted)?.into_inner();
                if data.len() != 1 {
                    return Err(error::Visibility::DataLength(data.len()));
                }
                match data.first() {
                    Some(0x00) => Ok(Visibility::NotVisible),
                    Some(0xff) => Ok(Visibility::Visible),
                    Some(x) => Err(error::Visibility::InvalidState(*x)),
                    None => unreachable!(),
                }
            },
        )(input)
    }

    fn host_extenders(input: &[u8]) -> IResult<&[u8], Vec<HostExtenderRef>> {
        preceded(
            tuple((
                new_line::parse,
                tag("[Host Extender Info]"),
                new_line::parse,
            )),
            many0(host_extender_ref),
        )(input)
    }

    fn host_extender_ref(input: &[u8]) -> IResult<&[u8], HostExtenderRef> {
        map(
            tuple((
                terminated(hex_int_32::parse, tag("=")),
                terminated(guid::parse, tag(";")),
                terminated(lib_name, tag(";")),
                terminated(hex_int_32::parse, new_line::parse),
            )),
            |(index, guid, lib, creation_flags)| HostExtenderRef {
                index,
                guid,
                lib,
                creation_flags,
            },
        )(input)
    }

    // TODO: This probably needs to take account of the MBCS requirment
    fn lib_name(input: &[u8]) -> IResult<&[u8], LibName> {
        map_res(take_while(|c| c > 0x20 && c != 0x3b), |s: &[u8]| {
            String::from_utf8(s.to_vec())
        })(input)
    }

    fn workspace(input: &[u8]) -> IResult<&[u8], Vec<WindowRecord>> {
        preceded(
            tuple((new_line::parse, tag("[Workspace]"), new_line::parse)),
            many0(window_record),
        )(input)
    }

    fn window_record(input: &[u8]) -> IResult<&[u8], WindowRecord> {
        map(
            tuple((
                terminated(module_identifier::parse, tag("=")),
                project_window,
                opt(preceded(tag(", "), project_window)),
                new_line::parse,
            )),
            |(module, code, designer, _)| WindowRecord {
                module,
                code,
                designer,
            },
        )(input)
    }

    fn project_window(input: &[u8]) -> IResult<&[u8], Window> {
        map(
            tuple((window_dim, window_dim, window_dim, window_dim, window_state)),
            |(left, top, right, bottom, state)| Window {
                left,
                top,
                right,
                bottom,
                state,
            },
        )(input)
    }

    fn window_dim(input: &[u8]) -> IResult<&[u8], int_32::Int32> {
        terminated(int_32::parse, tag(", "))(input)
    }

    fn window_state(input: &[u8]) -> IResult<&[u8], WindowState> {
        map(one_of("CZI"), |c| match c {
            'C' => WindowState::Closed,
            'Z' => WindowState::Zoomed,
            'I' => WindowState::Minimized,
            _ => unreachable!(),
        })(input)
    }
}
