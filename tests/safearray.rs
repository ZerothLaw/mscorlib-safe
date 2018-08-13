extern crate mscorlib_safe;
extern crate mscorlib_sys;
extern crate winapi;

use std::ptr;

use winapi::ctypes::c_void;
use winapi::shared::wtypes::{VT_I2, VARTYPE};
use winapi::um::oaidl::LPSAFEARRAY;
use winapi::um::oaidl::SAFEARRAYBOUND;
use winapi::um::oaidl::IDispatch;

use mscorlib_sys::system::reflection::_Type;

use mscorlib_safe::new_variant::{Variant, UInt, Int};
use mscorlib_safe::new_safearray::{RSafeArray, SafeArrayCreate, SafeArrayDestroy, SafeArrayPutElement, SafeArrayGetVartype, SafeArrayGetElement};

#[test]
fn from_rust_safearray_to_low_level() {
    let mut v: Vec<i16> = Vec::new();
    for ix in 0i16..100i16 {
        v.push(ix)
    }
    let rsa: RSafeArray<i16> = RSafeArray::Shorts(v);
    let ovt = rsa.vartype();
    let psa = LPSAFEARRAY::from(rsa);

    let mut vt:VARTYPE = 0;
    let hr = unsafe {
        SafeArrayGetVartype(psa, &mut vt)
    };
    assert_eq!(hr, 0);
    assert_eq!(ovt, vt as u32);

    let mut val: i16 = 0;
    let hr = unsafe {
        SafeArrayGetElement(psa, &10, &mut val as *mut _ as *mut c_void)
    };

    assert_eq!(hr, 0);
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

    let rsa: RSafeArray<u16> = RSafeArray::from(psa);
    assert_eq!(rsa.len(), 9);
    assert_eq!(rsa.vartype(), VT_I2);
    if let RSafeArray::Shorts(array) = rsa {
        assert_eq!(array[3], 3)
    } else {
        panic!("Incorrect type")   
    };
}

#[test]
fn test_i32() {
    let mut vc = Vec::new();
    for ix in 0..100i32 {
        vc.push(ix);
    }
    let rsa: RSafeArray<i32> = RSafeArray::from(vc);
    if let RSafeArray::Longs(array) = rsa {
        assert_eq!(array[50], 50);
    }
    else {
        panic!("Incorrect type");
    }
}

#[test]
fn test_f32() {
    let mut vc = Vec::new();
    let mut val = 0.0f32;
    for _ix in 0..100 {
        vc.push(val);
        val += 1.0f32;
    }
    let rsa: RSafeArray<f32> = RSafeArray::from(vc);
    if let RSafeArray::Floats(array) = rsa {
        assert_eq!(array[50], 50.0);
    }
    else {
        panic!("Incorrect type");
    }
}

#[test]
fn test_f64() {
    let mut vc = Vec::new();
    let mut val = 0.0f64;
    for _ix in 0..100 {
        vc.push(val);
        val += 1.0f64;
    }
    let rsa: RSafeArray<f64> = RSafeArray::from(vc);
    if let RSafeArray::Doubles(array) = rsa {
        assert_eq!(array[50], 50.0);
    }
    else {
        panic!("Incorrect type");
    }
}

#[test]
fn test_str() {
    let mut vc = Vec::new();
    let mut val = 0.0f64;
    for _ix in 0..100 {
        vc.push(val.to_string());
        val += 1.0f64;
    }
    let rsa: RSafeArray<String> = RSafeArray::from(vc);
    if let RSafeArray::BStrings(array) = rsa {
        assert_eq!(array[50] , "50");
    }
    else {
        panic!("Incorrect type");
    }
}

#[test]
fn test_idispatch() {
    let mut vc = Vec::new();
    for _ix in 0..100 {
        let p: *mut IDispatch = ptr::null_mut();
        vc.push(p);
    }
    let rsa: RSafeArray<_Type> = RSafeArray::from(vc);
    if let RSafeArray::Dispatchs(array, _) = rsa {
        assert!(array[50].is_null());
    }
    else {
        panic!("Incorrect type");
    }
}

#[test]
fn test_bool() {
    let mut vc = Vec::new();
    for ix in 0..100 {
        vc.push(ix % 2 == 0);
    }
    let rsa: RSafeArray<bool> = RSafeArray::from(vc);
    if let RSafeArray::Bools(array) = rsa {
        assert_eq!(array[50], true);
    }
    else {
        panic!("Incorrect type");
    }
}

#[test]
fn test_variant() {
    let mut vc = Vec::new();
    for ix in 0..100 {
        let vt = Variant::from(ix % 2 == 0);
        vc.push(vt);
    }
    let rsa: RSafeArray<Variant> = RSafeArray::from(vc);
    if let RSafeArray::Variants(array) = rsa {
        assert_eq!(array[50], Variant::Bool(true));
    }
    else {
        panic!("Incorrect type");
    }
}

#[test]
fn test_i8() {
    let mut vc = Vec::new();
    for ix in 0..100i8 {
        vc.push(ix);
    }
    let rsa:RSafeArray<i8> = RSafeArray::from(vc);
    if let RSafeArray::Chars(array) = rsa {
        assert_eq!(array[50], 50i8);
    }
    else {
        panic!("Incorrect type");
    }
}

#[test]
fn test_u8() {
    let mut vc = Vec::new();
    for ix in 0..100u8 {
        vc.push(ix);
    }
    let rsa:RSafeArray<u8> = RSafeArray::from(vc);
    if let RSafeArray::UChars(array) = rsa {
        assert_eq!(array[50], 50u8);
    }
    else {
        panic!("Incorrect type");
    }
}

#[test]
fn test_u16() {
    let mut vc = Vec::new();
    for ix in 0..100u16 {
        vc.push(ix);
    }
    let rsa: RSafeArray<u16> = RSafeArray::from(vc);
    if let RSafeArray::UShorts(array) = rsa {
        assert_eq!(array[50], 50u16);
    }
    else {
        panic!("Incorrect type");
    }
}

#[test]
fn test_u32() {
    let mut vc = Vec::new();
    for ix in 0..100u32 {
        vc.push(ix);
    }
    let rsa: RSafeArray<u32> = RSafeArray::from(vc);
    if let RSafeArray::ULongs(array) = rsa {
        assert_eq!(array[50], 50u32);
    }
    else {
        panic!("Incorrect type");
    }
}

#[test]
fn test_int() {
    let mut vc = Vec::new();
    let val = 1i32 << 30;
    for ix in val..(val+100){
        vc.push(Int(ix));
    }
    let rsa: RSafeArray<i64> = RSafeArray::from(vc);
    if let RSafeArray::Ints(array) = rsa {
        assert_eq!(array[50], Int(1073741874i32));
    }
    else {
        panic!("Incorrect type");
    }
}

#[test]
fn test_uint() {
    let mut vc = Vec::new();
    let val = 1u32 << 30;
    for ix in val..(val+100){
        vc.push(UInt(ix));
    }
    let rsa: RSafeArray<u64> = RSafeArray::from(vc);
    if let RSafeArray::UInts(array) = rsa {
        assert_eq!(array[50], UInt(1073741874u32));
    }
    else {
        panic!("Incorrect type");
    }
}
