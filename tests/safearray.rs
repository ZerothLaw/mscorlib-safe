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

use mscorlib_safe::new_variant::{Variant, Primitive};
use mscorlib_safe::new_safearray::{RSafeArray, SafeArrayCreate, SafeArrayDestroy, SafeArrayPutElement, SafeArrayGetLBound, SafeArrayGetUBound,SafeArrayGetVartype, SafeArrayGetElement};

#[test]
fn from_rust_safearray_to_low_level() {
    let mut v: Vec<i16> = Vec::new();
    for ix in 0i16..100i16 {
        v.push(ix)
    }
    let rsa: RSafeArray<i16> = RSafeArray::I16(v);
    let ovt = rsa.vartype();
    let psa = LPSAFEARRAY::from(rsa);

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

    let rsa: RSafeArray<u16> = RSafeArray::from(psa);
    assert_eq!(rsa.len(), 9);
    assert_eq!(rsa.vartype(), VT_I2);
    if let RSafeArray::I16(array) = rsa {
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
    if let RSafeArray::I32(array) = rsa {
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
    for ix in 0..100 {
        vc.push(val);
        val += 1.0f32;
    }
    let rsa: RSafeArray<f32> = RSafeArray::from(vc);
    if let RSafeArray::F32(array) = rsa {
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
    for ix in 0..100 {
        vc.push(val);
        val += 1.0f64;
    }
    let rsa: RSafeArray<f64> = RSafeArray::from(vc);
    if let RSafeArray::F64(array) = rsa {
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
    for ix in 0..100 {
        vc.push(val.to_string());
        val += 1.0f64;
    }
    let rsa: RSafeArray<String> = RSafeArray::from(vc);
    if let RSafeArray::BString(array) = rsa {
        assert_eq!(array[50] , "50");
    }
    else {
        panic!("Incorrect type");
    }
}

#[test]
fn test_idispatch() {
    let mut vc = Vec::new();
    for ix in 0..100 {
        let p: *mut IDispatch = ptr::null_mut();
        vc.push(p);
    }
    let rsa: RSafeArray<_Type> = RSafeArray::from(vc);
    if let RSafeArray::Dispatch(array, _) = rsa {
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
    if let RSafeArray::Bool(array) = rsa {
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
    if let RSafeArray::Variant(array) = rsa {
        assert_eq!(array[50], Variant::VariantPrimitive(Primitive::Bool(true)));
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
    if let RSafeArray::SChar(array) = rsa {
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
    if let RSafeArray::UChar(array) = rsa {
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
    if let RSafeArray::UShort(array) = rsa {
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
    if let RSafeArray::ULong(array) = rsa {
        assert_eq!(array[50], 50u32);
    }
    else {
        panic!("Incorrect type");
    }
}

#[cfg(target_arch="x86_64")]
#[test]
fn test_int() {
    let mut vc = Vec::new();
    let val = 1i64 << 32;
    for ix in val..(val+100){
        vc.push(ix);
    }
    let rsa: RSafeArray<i64> = RSafeArray::from(vc);
    if let RSafeArray::Int(array) = rsa {
        assert_eq!(array[50], 4294967346i64);
    }
    else {
        panic!("Incorrect type");
    }
}

#[cfg(target_arch="x86_64")]
#[test]
fn test_uint() {
    let mut vc = Vec::new();
    let val = 1u64 << 32;
    for ix in val..(val+100){
        vc.push(ix);
    }
    let rsa: RSafeArray<u64> = RSafeArray::from(vc);
    if let RSafeArray::UInt(array) = rsa {
        assert_eq!(array[50], 4294967346u64);
    }
    else {
        panic!("Incorrect type");
    }
}

#[cfg(target_arch="x86")]
#[test]
fn test_int() {
    let mut vc = Vec::new();
    let val = 1i32 << 32;
    for ix in val..(val+100){
        vc.push(ix);
    }
    let rsa: RSafeArray<i32> = RSafeArray::from(vc);
    if let RSafeArray::Int(array) = rsa {
        assert_eq!(array[50], 4294967346i32);
    }
    else {
        panic!("Incorrect type");
    }
}

#[cfg(target_arch="x86")]
#[test]
fn test_uint() {
    let mut vc = Vec::new();
    let val = 1u32 << 32;
    for ix in val..(val+100){
        vc.push(ix);
    }
    let rsa: RSafeArray<i32> = RSafeArray::from(vc);
    if let RSafeArray::UInt(array) = rsa {
        assert_eq!(array[50], 4294967346u32);
    }
    else {
        panic!("Incorrect type");
    }
}