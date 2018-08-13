extern crate winapi;

extern crate mscorlib_sys;
extern crate mscorlib_safe;

#[macro_use]
extern crate mscorlib_safe_derive;

use std::ptr;

use winapi::um::oaidl::{LPSAFEARRAY};

use mscorlib_sys::system::reflection::_Type;
use mscorlib_safe::PtrContainer;
use mscorlib_safe::new_variant::Variant;
use mscorlib_safe::new_safearray::RSafeArray;

use mscorlib_sys::system::reflection::_TypeVtbl;

#[derive(Debug,PtrContainer)]
struct Fit {
    ptr: *mut _Type
}

fn main() {
    let mut vc = Vec::new();
    vc.push(100u32);

    let rsa: RSafeArray<u32> = RSafeArray::from(vc);
    let psa = LPSAFEARRAY::from(rsa);
}