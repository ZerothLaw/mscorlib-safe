#![feature(never_type)]

#[macro_use] extern crate failure;

extern crate winapi;

extern crate mscorlib_sys;

mod bstring;
#[macro_use]pub mod macros;
mod params;
mod primitives;
mod result;
mod safearray;
mod struct_wrappers;
#[macro_use]mod variant;
mod wrappers;

pub mod new_variant;
pub mod new_safearray;

pub use bstring::*;
pub use primitives::*;
pub use result::*;
pub use safearray::*;
pub use variant::*;
pub use wrappers::*;

