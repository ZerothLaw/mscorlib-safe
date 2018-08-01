#![feature(never_type)]

#[macro_use] extern crate failure;

extern crate winapi;

extern crate mscorlib_sys;

mod bstring;
#[macro_use]pub mod macros;
mod params;
mod result;
mod safearray;
mod struct_wrappers;
#[macro_use]mod variant;
mod wrappers;

pub use bstring::*;
pub use result::*;
pub use safearray::*;
pub use variant::*;
pub use wrappers::*;
