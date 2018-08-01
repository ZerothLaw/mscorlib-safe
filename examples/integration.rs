extern crate winapi;

extern crate mscorlib_sys;
extern crate mscorlib_safe;

#[macro_use]
extern crate mscorlib_safe_derive;

use std::ptr;

use winapi::um::oaidl::{LPSAFEARRAY};

use mscorlib_sys::system::reflection::_Type;
use mscorlib_safe::PtrContainer;
use mscorlib_safe::SafeArray;
use mscorlib_safe::{WrappedDispatch, PhantomDispatch};

#[derive(PtrContainer)]
struct Fit {
    ptr: *mut _Type
}

fn main() {
    let mut vc = Vec::new();
    vc.push(Fit{ptr: ptr::null_mut()});
    vc.push(Fit{ptr: ptr::null_mut()});
    vc.push(Fit{ptr: ptr::null_mut()});
    let w: SafeArray<WrappedDispatch, Fit, PhantomDispatch, _Type, String> = SafeArray::SafeUnknown(vc);
    let s = LPSAFEARRAY::from(w);
    println!("Yay: {:p}", s);
}