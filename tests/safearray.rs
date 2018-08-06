extern crate mscorlib_safe;
extern crate mscorlib_sys;
extern crate winapi;

use std::ptr;

use winapi::ctypes::c_void;
use winapi::shared::wtypes::VARTYPE;
use winapi::um::oaidl::SAFEARRAY;

use mscorlib_safe::{UnknownSafeArray, DispatchSafeArray, SafeArray, PrimitiveSafeArray, Primitives, Primitive, SafeArrayGetElement, SafeArrayGetVartype};

#[test]
fn from_rust_safearray_to_low_level() {
    let mut v:Vec<Box<Primitive<Target=i16>>> = Vec::new();
    for ix in 0i16..100i16 {
        v.push(Box::new(ix))
    }

    let sa: PrimitiveSafeArray<i16> = SafeArray::SafePrimitive(v);

    /*impl<TD, TU, PD, PU> From<SafeArray<TD, TU, PD, PU, String>> for LPSAFEARRAY 
    where TD: PtrContainer<PD>, 
          TU: PtrContainer<PU>,
          PD: Deref<Target=IDispatch>,
          PU: Deref<Target=IUnknown>,*/

    let psa: *mut SAFEARRAY = From::from(sa);
    unsafe {
        let psab = (*psa).rgsabound;
        assert_eq!(psab[0].cElements, 100);
        let ix = [0];
        let mut v = -1i16;
        assert_eq!(v, -1);
        let pv: *mut i16 = &mut v;
        let pv: *mut c_void = pv as *mut _;
        let mut vt: VARTYPE = 0;
        
        let hr = SafeArrayGetVartype(psa, &mut vt);
        println!("vt: {}", vt);
        println!("pv: {:p}", pv);
        let hr = SafeArrayGetElement(psa, ix.as_ptr() , pv);
        println!("0x{:x}", hr);
        assert_eq!(hr, 0);
        println!("{:p}", pv);
        println!("{:?}", *(pv as *mut _ as *mut i16));
        //println!("pv = {:p}", *pv);
        assert_eq!(v, 20);
    }

}