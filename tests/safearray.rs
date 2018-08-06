extern crate mscorlib_safe;
extern crate mscorlib_sys;
extern crate winapi;

use std::ptr;

use winapi::ctypes::c_void;
use winapi::shared::wtypes::{VT_I2, VARTYPE};
use winapi::um::oaidl::LPSAFEARRAY;
use winapi::um::oaidl::SAFEARRAYBOUND;

//use mscorlib_safe::new_variant::Variant;
use mscorlib_safe::new_safearray::{RSafeArray, SafeArrayCreate, SafeArrayDestroy, SafeArrayPutElement, SafeArrayGetLBound, SafeArrayGetUBound,SafeArrayGetVartype, SafeArrayGetElement};

#[test]
fn from_rust_safearray_to_low_level() {
    let mut v: Vec<i16> = Vec::new();
    for ix in 0i16..100i16 {
        v.push(ix)
    }
    let mut rsa = RSafeArray::I16(v);
    let ovt = rsa.vartype();
    let mut psa = LPSAFEARRAY::from(rsa);

    let mut vt:VARTYPE = 0;
    let hr = unsafe {
        SafeArrayGetVartype(psa, &mut vt)
    };
    //println!("hr: 0x{:x}", hr);
    assert_eq!(ovt, vt as u32);

    let mut val: i16 = 0;
    let hr = unsafe {
        SafeArrayGetElement(psa, &10, &mut val as *mut _ as *mut c_void)
    };
    //println!("hr: 0x{:x}", hr);
    assert_eq!(val, 10);

    unsafe {SafeArrayDestroy(psa)};
}

#[test]
fn from_low_level_to_rust_safearray(){
    let mut sab = SAFEARRAYBOUND{cElements: 10, lLbound: 0};
    let psa = unsafe{SafeArrayCreate(VT_I2 as u16, 1, &mut sab)};
    for ix in 0..10 {
        let mut val = ix;
        let hr = unsafe {
            SafeArrayPutElement(psa, &ix, &mut val as *mut _ as *mut c_void)
        };
        assert_eq!(hr, 0);
    }

    let rsa = RSafeArray::from(psa);
    assert_eq!(rsa.len(), 9);
    assert_eq!(rsa.vartype(), VT_I2);
    if let RSafeArray::I16(array) = rsa {
        assert_eq!(array[3], 3)
    };
}