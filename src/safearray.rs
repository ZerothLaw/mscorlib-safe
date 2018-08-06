use std::fmt::Debug;
use std::marker::PhantomData;
use std::ops::{Deref, Index};
use std::ptr;


use winapi::ctypes::{c_long, c_void};
use winapi::shared::ntdef::{LONG, ULONG};
use winapi::shared::minwindef::UINT;
use winapi::shared::winerror::HRESULT;
use winapi::shared::wtypes::{VARTYPE, VT_BSTR, VT_DISPATCH, VT_UNKNOWN};

use winapi::um::oaidl::{LPSAFEARRAY, LPSAFEARRAYBOUND, SAFEARRAYBOUND, IDispatch};
use winapi::um::unknwnbase::IUnknown;

use bstring::BString;
use primitives::Primitive;
use wrappers::PtrContainer;
use variant::{WrappedDispatch, WrappedUnknown, PhantomDispatch, PhantomUnknown};

/*
STRUCT!{struct SAFEARRAYBOUND {
    cElements: ULONG,
    lLbound: LONG,
}}

STRUCT!{struct SAFEARRAY {
    cDims: USHORT,
    fFeatures: USHORT,
    cbElements: ULONG,
    cLocks: ULONG,
    pvData: PVOID,
    rgsabound: [SAFEARRAYBOUND; 1],
}}
*/

extern "system" {
    pub fn SafeArrayCreate(vt: VARTYPE, cDims: UINT, rgsabound: LPSAFEARRAYBOUND) -> LPSAFEARRAY;
	pub fn SafeArrayDestroy(safe: LPSAFEARRAY)->HRESULT;
    
    pub fn SafeArrayGetDim(psa: LPSAFEARRAY) -> UINT;
	
    pub fn SafeArrayGetElement(psa: LPSAFEARRAY, rgIndices: *const c_long, pv: *mut c_void) -> HRESULT;
    pub fn SafeArrayGetElemSize(psa: LPSAFEARRAY) -> UINT;
    
    pub fn SafeArrayGetLBound(psa: LPSAFEARRAY, nDim: UINT, plUbound: *mut c_long)->HRESULT;
    pub fn SafeArrayGetUBound(psa: LPSAFEARRAY, nDim: UINT, plUbound: *mut c_long)->HRESULT;
    
    pub fn SafeArrayGetVartype(psa: LPSAFEARRAY, pvt: *mut VARTYPE) -> HRESULT;

    pub fn SafeArrayLock(psa: LPSAFEARRAY) -> HRESULT;
	pub fn SafeArrayUnlock(psa: LPSAFEARRAY) -> HRESULT;
    
    pub fn SafeArrayPutElement(psa: LPSAFEARRAY, rgIndices: *const c_long, pv: *mut c_void) -> HRESULT;
}

//newtype it
pub enum SafeArray<TDispatch, TUnknown, PDispatch, PUnknown, B, TPrimitive> 
    where TDispatch: PtrContainer<PDispatch>, 
          TUnknown: PtrContainer<PUnknown>,
          PDispatch: Deref<Target=IDispatch>,
          PUnknown: Deref<Target=IUnknown>,
          B: From<BString>
{
    SafeDispatch(Vec<TDispatch>), 
    SafeUnknown(Vec<TUnknown>), 
    SafeBstr(Vec<B>),
    SafePrimitive(Vec<Box<Primitive<Target=TPrimitive>>>),
    Empty,
    Phantom(PhantomData<PDispatch>, PhantomData<PUnknown>, !)
}

pub use SafeArray::*;

pub type DispatchSafeArray<TDispatch, PDispatch> = SafeArray<TDispatch, WrappedUnknown, PDispatch, PhantomUnknown, String, i16>;
pub type UnknownSafeArray<TUnknown, PUnknown> = SafeArray<WrappedDispatch, TUnknown, PhantomDispatch, PUnknown, String, i16>;
pub type StringSafeArray = SafeArray<WrappedDispatch, WrappedUnknown, PhantomDispatch, PhantomUnknown, String, i16>;
pub type PrimitiveSafeArray<TPrimitive> = SafeArray<WrappedDispatch, WrappedUnknown, PhantomDispatch, PhantomUnknown, String, TPrimitive>;
// impl<TD, TU, PD, PU> SafeArray<TD, TU, PD, PU, String> 
//     where TD: PtrContainer<PD>, 
//           TU: PtrContainer<PU>,
//           PD: Deref<Target=IDispatch>,
//           PU: Deref<Target=IUnknown>,
// {
//     fn new() -> SafeArray<TD, TU, PD, PU, String> {

//     }
// }

impl<TD, TU, PD, PU, TP> From<SafeArray<TD, TU, PD, PU, String, TP>> for LPSAFEARRAY 
    where TD: PtrContainer<PD>, 
          TU: PtrContainer<PU>,
          PD: Deref<Target=IDispatch>,
          PU: Deref<Target=IUnknown>,
          TP: Debug
{
    fn from(array: SafeArray<TD, TU, PD, PU, String, TP>) -> LPSAFEARRAY {
        let (vartype, sz) = match array {
            SafeDispatch(ref array) => (VT_DISPATCH, array.len()), 
            SafeUnknown(ref array) => (VT_UNKNOWN, array.len()), 
            SafeBstr(ref array) => (VT_BSTR, array.len()), 
            SafePrimitive(ref array) => {
                if array.len() > 0 {
                    let f = array.index(0);
                    (VARTYPE::from(f.prim_type()) as u32, array.len())
                }
                else { (0, 0) }
            },
            Empty | Phantom(_,_, _) => (0, 0),
        };
        println!("{:?}", (vartype, sz));
        let mut sab = SAFEARRAYBOUND{ cElements: sz as ULONG, lLbound: 0 as LONG};
        let psa = unsafe {
            SafeArrayCreate(vartype as u16, 1, &mut sab)
        };
        match array {
            SafeDispatch(array) => {
                for (ix,elem) in array.into_iter().enumerate() {
                    let ptr = elem.ptr_mut();
                    let hr = unsafe {
                        SafeArrayPutElement(psa, &(ix as i32), ptr as *mut _ as *mut c_void)
                    };
                    match hr {
                        0 => continue,
                        _ => panic!("Error: 0x{:x} occurred at index: {}", hr, ix)
                    }
                }
            }, 
            SafeUnknown(array) => {
                for (ix,elem) in array.into_iter().enumerate() {
                    let ptr = elem.ptr_mut();
                    let hr = unsafe {
                        SafeArrayPutElement(psa, &(ix as i32), ptr as *mut _ as *mut c_void)
                    };
                    match hr {
                        0 => continue,
                        _ => panic!("Error: 0x{:x} occurred at index: {}", hr, ix)
                    }
                }
            }, 
            SafeBstr(array) => {
                for (ix,elem) in array.into_iter().enumerate() {
                    let elem: BString = From::from(elem);
                    let hr = unsafe {
                        let ptr = elem.as_sys();
                        SafeArrayPutElement(psa, &(ix as i32), ptr as *mut _ as *mut c_void)
                    };
                    match hr {
                        0 => continue,
                        _ => panic!("Error: 0x{:x} occurred at index: {}", hr, ix)
                    }
                }
            },
            SafePrimitive(array) => {
                for (ix,elem) in array.into_iter().enumerate() {
                    let mut elem = elem.get();
                    println!("elem: {:?}", elem);
                    let hr = unsafe {
                        SafeArrayPutElement(psa, &(ix as i32), &mut elem as *mut _ as *mut c_void)
                    };
                    match hr {
                        0 => continue,
                        _ => panic!("Error: 0x{:x} occurred at index: {}", hr, ix)
                    }
                }
            }
            Empty => {},
            Phantom(_,_,_) => {unreachable!()}
        }
        psa
    }
}

impl<TD, TU, PD, PU, TP> From<LPSAFEARRAY> for SafeArray<TD, TU, PD, PU, String, TP> 
    where TD: PtrContainer<PD>, 
          TU: PtrContainer<PU>,
          PD: Deref<Target=IDispatch>,
          PU: Deref<Target=IUnknown>,
{
    fn from(psa: LPSAFEARRAY) -> SafeArray<TD, TU, PD, PU, String, TP> {
        let mut vd = Vec::new();
        let mut vb = Vec::new();
        let mut vu = Vec::new();
        let mut vartype: VARTYPE = 0;
        let hr = unsafe {
            SafeArrayGetVartype(psa, &mut vartype)
        };
        let imm_vartype = vartype;
        match hr {
            0 => {}, 
            _ => panic!("Error occurred with SafeArrayGetVartype call: 0x{:x}", hr)
        };
        let dims = unsafe {
            SafeArrayGetDim(psa)
        };
        let mut dimensions: Vec<(i32, i32)> = Vec::new();
        for dim in 1..dims {
            let mut lower = 0;
            let hr = unsafe {
                SafeArrayGetLBound(psa, dim, &mut lower)
            };
            match hr {
                0 => {}, 
                _ => panic!("Error occurred with SafeArrayGetLBound call: 0x{:x}", hr)
            };
            let mut upper = 0;
            let hr = unsafe {
                SafeArrayGetUBound(psa, dim, &mut upper)
            };
            match hr {
                0 => {}, 
                _ => panic!("Error occurred with SafeArrayGetUBound call: 0x{:x}", hr)
            };
            dimensions.push((lower, upper));
        };
        let mut slice: Vec<i32> = dimensions.iter().map(|item| item.0 ).collect();
        let beg_slice: Vec<i32> = dimensions.iter().map(|item| item.0 ).collect();
        let end_slice: Vec<i32> = dimensions.iter().map(|item| item.1 ).collect();
        while slice != end_slice {
            let pv: *mut c_void = ptr::null_mut();
            let hr = unsafe {
                SafeArrayGetElement(psa, slice[..].as_ptr(), pv )
            };
            match hr {
                0 => {
                    match imm_vartype as u32 {
                        VT_DISPATCH => {
                            let wrapper = TD::from(pv as *mut _ as *mut PD);
                            vd.push(wrapper);
                        }, 
                        VT_UNKNOWN => {
                            let wrapper = TU::from(pv as *mut _ as *mut PU);
                            vu.push(wrapper);
                        }, 
                        VT_BSTR => {
                            let bs = BString::from_ptr_safe(pv as *mut _ as *mut u16);

                            vb.push(bs.to_string());
                        }, 
                        _ => {}
                    }
                   
                    //iterate the indices
                    let mut ix = slice.len()-1;
                    let copy_slice = slice.clone();
                    for val in copy_slice.iter().rev() {
                        let val = val + 1;
                        if val > end_slice[ix] {
                            slice[ix] = beg_slice[ix];
                            ix -= 1;
                            continue;
                        }
                        else {
                            slice[ix] = val;
                            break;
                        }
                    }
                },
                _ => panic!("Error(0x{:x}) accessing element at {:?}", hr, slice)
            }
        }
        match imm_vartype as u32 {
            VT_BSTR => SafeBstr(vb), 
            VT_DISPATCH => SafeDispatch(vd), 
            VT_UNKNOWN => SafeUnknown(vu),
            _ => Empty
        }
    }
}
