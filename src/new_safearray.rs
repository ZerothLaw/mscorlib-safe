use std::fmt::Debug;
use std::marker::PhantomData;
use std::mem;
use std::ops::Deref;
use std::ptr;

use winapi::ctypes::{c_long, c_void};

use winapi::shared::minwindef::{UINT, ULONG};
use winapi::shared::winerror::HRESULT;
use winapi::shared::wtypes::{CY, VARENUM, VARTYPE, VT_BOOL,    VT_BSTR, 
                             VT_CY,   VT_DATE, VT_DECIMAL, VT_DISPATCH,                              VT_INT,  VT_I1,  VT_I2,   VT_I4,   VT_R4,       
                             VT_R8,   VT_UINT, VT_UNKNOWN, VT_UI1,  
                             VT_UI2,  VT_UI4,  VT_VARIANT};
use winapi::shared::wtypes::{BSTR, VARIANT_BOOL};

use winapi::um::oaidl::{IDispatch, LPSAFEARRAYBOUND, SAFEARRAY, SAFEARRAYBOUND, VARIANT};
use winapi::um::unknwnbase::IUnknown;
use rust_decimal::Decimal;

use new_variant::{Currency, Date, Int, UInt, Variant, build_c_decimal};

use wrappers::PtrContainer;

use bstring;

/*{ 
if let RSafeArray::Unknown(array) = RSafeArray::from(pmodules) {
    array.into_iter().map(|item| {
        let trans_item = unsafe { mem::transmute::<*mut IUnknown, *mut _Module>(item)};
        M::from(trans_item)
    }).collect()
}
else {
    Vec::new()
}*/

extern "system" {
    pub fn SafeArrayCreate(vt: VARTYPE, cDims: UINT, rgsabound: LPSAFEARRAYBOUND) -> LPSAFEARRAY;
	pub fn SafeArrayDestroy(safe: LPSAFEARRAY)->HRESULT;
    
    pub fn SafeArrayGetDim(psa: LPSAFEARRAY) -> UINT;
	
    pub fn SafeArrayGetElement(psa: LPSAFEARRAY, rgIndices: *const c_long, pv: *mut c_void) -> HRESULT;
    pub fn SafeArrayGetElemSize(psa: LPSAFEARRAY) -> UINT;
    
    pub fn SafeArrayGetLBound(psa: LPSAFEARRAY, nDim: UINT, plLbound: *mut c_long)->HRESULT;
    pub fn SafeArrayGetUBound(psa: LPSAFEARRAY, nDim: UINT, plUbound: *mut c_long)->HRESULT;
    
    pub fn SafeArrayGetVartype(psa: LPSAFEARRAY, pvt: *mut VARTYPE) -> HRESULT;

    pub fn SafeArrayLock(psa: LPSAFEARRAY) -> HRESULT;
	pub fn SafeArrayUnlock(psa: LPSAFEARRAY) -> HRESULT;
    
    pub fn SafeArrayPutElement(psa: LPSAFEARRAY, rgIndices: *const c_long, pv: *mut c_void) -> HRESULT;
}

pub use winapi::um::oaidl::LPSAFEARRAY;

#[derive(Debug, Clone, PartialEq, PartialOrd)]
pub enum RSafeArray<P=i32> {
    Shorts(Vec<i16>), //VT_I2
    Longs(Vec<i32>), //VT_I4,
    Floats(Vec<f32>), //VT_R4, 
    Doubles(Vec<f64>), //VT_R8, 
    Currencies(Vec<Currency>), //VT_CY, 
    Dates(Vec<Date>),//VT_DATE,
    BStrings(Vec<String>), //VT_BSTR,
    Dispatchs(Vec<*mut IDispatch>, Option<P>), //VT_DISPATCH, 
    Bools(Vec<bool>), //VT_BOOL, need to translate between rust bool and VARIANT_BOOL,
    Variants(Vec<Variant>), //VT_VARIANT,
    Unknowns(Vec<*mut IUnknown>, Option<P>), //VT_UNKNOWN, 
    Decimals(Vec<Decimal>),//VT_DECIMAL,
    //VT_RECORD,
    Chars(Vec<i8>), //VT_I1, 
    UChars(Vec<u8>), //VT_UI1, 
    UShorts(Vec<u16>), //VT_UI2, 
    ULongs(Vec<u32>), //VT_UI4,
    Ints(Vec<Int>), //VT_INT, 
    UInts(Vec<UInt>), //VT_UINT,
}

impl<T> RSafeArray<T> {
    pub fn len(&self) -> usize {
        match self {
            RSafeArray::Shorts(inner) => inner.len(),
            RSafeArray::Longs(inner) => inner.len(),
            RSafeArray::Floats(inner) => inner.len(),
            RSafeArray::Doubles(inner) => inner.len(),
            RSafeArray::Currencies(inner) => inner.len(),
            RSafeArray::Dates(inner) => inner.len(),
            RSafeArray::BStrings(inner) => inner.len(), 
            RSafeArray::Dispatchs(inner, _) => inner.len(), 
            RSafeArray::Bools(inner) => inner.len(), 
            RSafeArray::Variants(inner) => inner.len(), 
            RSafeArray::Unknowns(inner, _) => inner.len(), 
            RSafeArray::Decimals(inner) => inner.len(),
            RSafeArray::Chars(inner) => inner.len(), 
            RSafeArray::UChars(inner) => inner.len(),
            RSafeArray::UShorts(inner) => inner.len(), 
            RSafeArray::ULongs(inner) => inner.len(), 
            RSafeArray::Ints(inner) => inner.len(), 
            RSafeArray::UInts(inner) => inner.len(),
        }
    }
    pub fn vartype(&self) -> VARENUM {
        match self { 
            RSafeArray::Shorts(_) => VT_I2, 
            RSafeArray::Longs(_) => VT_I4, 
            RSafeArray::Floats(_) => VT_R4, 
            RSafeArray::Doubles(_) => VT_R8, 
            RSafeArray::Currencies(_) => VT_CY,
            RSafeArray::Dates(_) => VT_DATE, 
            RSafeArray::BStrings(_) => VT_BSTR,
            RSafeArray::Dispatchs(_, _) => VT_DISPATCH, 
            RSafeArray::Bools(_) => VT_BOOL, 
            RSafeArray::Variants(_) => VT_VARIANT, 
            RSafeArray::Unknowns(_, _) => VT_UNKNOWN, 
            RSafeArray::Decimals(_) => VT_DECIMAL,
            RSafeArray::Chars(_) => VT_I1, 
            RSafeArray::UChars(_) => VT_UI1, 
            RSafeArray::UShorts(_) => VT_UI2, 
            RSafeArray::ULongs(_) => VT_UI4, 
            RSafeArray::Ints(_) => VT_INT, 
            RSafeArray::UInts(_) => VT_UINT,
        }
    }

    pub fn from_vec_dispatch<TOut: Deref<Target=IDispatch>>(vec: Vec<T>) -> RSafeArray<T> 
        where T: PtrContainer<TOut>
    {
        let stripped_vec: Vec<*mut IDispatch> = vec.into_iter().map(|item|{
            let ptr = item.ptr_mut();
            let id = unsafe { mem::transmute::<*mut TOut, *mut IDispatch>(ptr)};
            id
        }).collect();
        RSafeArray::Dispatchs(stripped_vec, None)
    }

    pub fn from_vec_unknown<TOut: Deref<Target=IUnknown>>(vec: Vec<T>) -> RSafeArray<T>
        where T: PtrContainer<TOut> 
    {
        let stripped_vec: Vec<*mut IUnknown> = vec.into_iter().map(|item|{
            let ptr = item.ptr_mut();
            let id = unsafe { mem::transmute::<*mut TOut, *mut IUnknown>(ptr)};
            id
        }).collect();
        RSafeArray::Unknowns(stripped_vec, None)
    }
}

macro_rules! FROM_IMPLS {
    (@branch, Dispatchs, $var:ident) => {
        RSafeArray::Dispatchs($var, None)
    };
    (@branch, Unknowns, $var:ident) => {
        RSafeArray::Unknowns($var, None)
    };
    (@branch, $enum_name:ident, $var:ident) => {
        RSafeArray::$enum_name($var)
    };
    ($($(#[$attrs:meta])* {$enum_name:ident, $v_type:ty})*) => {
        $(
            $(#[$attrs])*
            impl<T> From<Vec<$v_type>> for RSafeArray<T> {
                fn from(in_vc: Vec<$v_type>) -> RSafeArray<T> {
                    FROM_IMPLS!(@branch, $enum_name, in_vc)
                }
            }
        )*
    };
}

FROM_IMPLS!{
    {Shorts, i16}
    {Longs, i32}
    {Floats, f32}
    {Doubles, f64}
    {BStrings, String}
    {Dispatchs, *mut IDispatch}
    {Bools, bool}
    {Variants, Variant}
    {Unknowns, *mut IUnknown}
    {Chars, i8}
    {UChars, u8}
    {UShorts, u16}
    {ULongs, u32}
    {Ints, Int}
    {UInts, UInt}
}

struct SafeArrayDestructor {
    inner: *mut SAFEARRAY, 
    _marker: PhantomData<SAFEARRAY>
}

impl SafeArrayDestructor {
    fn new(p: *mut SAFEARRAY) -> SafeArrayDestructor {
        assert!(!p.is_null());
        SafeArrayDestructor{
            inner: p, 
            _marker: PhantomData
        }
    }
}

impl Drop for SafeArrayDestructor {
    fn drop(&mut self)  {
        if self.inner.is_null(){
            return;
        }
        unsafe {
            SafeArrayDestroy(self.inner)
        };
        self.inner = ptr::null_mut();
    }
}

impl<T> From<*mut SAFEARRAY> for RSafeArray<T> {
    fn from(psa: *mut SAFEARRAY) -> RSafeArray<T> {
        let _sad = SafeArrayDestructor::new(psa);
        let sa_dims = unsafe {SafeArrayGetDim(psa)};
        assert!(sa_dims > 0); //ensure we aren't dealing with a dimensionless safearr
        let vt = unsafe {
            let mut vt: VARTYPE = 0;
            let hr = SafeArrayGetVartype(psa, &mut vt);
            println!("hr = 0x{:x}", hr);
            vt
        };
        if sa_dims == 1 {
            let (lower_bound, upper_bound) = unsafe {
                let mut lower_bound: c_long = 0;
                let mut upper_bound: c_long = 0;
                let hr = SafeArrayGetLBound(psa, 1, &mut lower_bound);
                println!("L - hr = 0x{:x}", hr);
                let hr = SafeArrayGetUBound(psa, 1, &mut upper_bound);
                println!("U - hr = 0x{:x}", hr);
                (lower_bound, upper_bound)
            };
            match vt as u32 {
                VT_I2 => unsafe {
                    let mut vc = Vec::new();
                    for ix in lower_bound..upper_bound {
                        let mut i = 0i16;
                        let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void);
                        println!("loop {} - hr = 0x{:x}", ix, hr);
                        vc.push(i);
                    }
                    RSafeArray::Shorts(vc)
                },
                VT_I4 => unsafe {
                    let mut vc = Vec::new();
                    for ix in lower_bound..upper_bound {
                        let mut i = 0i32;
                        let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void);
                        println!("loop {} - hr = 0x{:x}", ix, hr);
                        vc.push(i);
                    }
                    RSafeArray::Longs(vc)
                },
                VT_R4 => unsafe {
                    let mut vc = Vec::new();
                    for ix in lower_bound..upper_bound {
                        let mut i:f32 = 0.0;
                        let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void);
                        println!("loop {} - hr = 0x{:x}", ix, hr);
                        vc.push(i);
                    }
                    RSafeArray::Floats(vc)
                },
                VT_R8 => unsafe {
                    let mut vc = Vec::new();
                    for ix in lower_bound..upper_bound {
                        let mut i:f64 = 0.0;
                        let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void);
                        println!("loop {} - hr = 0x{:x}", ix, hr);
                        vc.push(i);
                    }
                    RSafeArray::Doubles(vc)
                },
                VT_BSTR => unsafe {
                    let mut vc = Vec::new();
                    for ix in lower_bound..upper_bound {
                        let mut bs: BSTR = ptr::null_mut();
                        let hr = SafeArrayGetElement(psa, &ix, &mut bs as *mut _ as *mut c_void);
                        println!("loop {} - hr = 0x{:x}", ix, hr);
                        let s = bstring::BString::from_ptr_safe(bs).to_string();
                        vc.push(s)
                    }
                    RSafeArray::BStrings(vc)
                },
                VT_DISPATCH => unsafe {
                    let mut vc = Vec::new();
                    for ix in lower_bound..upper_bound {
                        let mut p: *mut IDispatch = ptr::null_mut();
                        let hr = SafeArrayGetElement(psa, &ix, &mut p as *mut _ as *mut c_void);
                        println!("loop {} - hr = 0x{:x}", ix, hr);
                        vc.push(p);
                    }
                    RSafeArray::Dispatchs(vc, None)
                },
                VT_BOOL => unsafe {
                    let mut vc = Vec::new();
                    for ix in lower_bound..upper_bound {
                        let mut vb: VARIANT_BOOL = 0;
                        let hr = SafeArrayGetElement(psa, &ix, &mut vb as *mut _ as *mut c_void);
                        println!("loop {} - hr = 0x{:x}", ix, hr);
                        vc.push(vb == -1);
                    }
                    RSafeArray::Bools(vc)
                },
                VT_VARIANT => unsafe {
                    let mut vc = Vec::new();
                    for ix in lower_bound..upper_bound {
                        let mut vt: *mut VARIANT = ptr::null_mut();
                        let hr = SafeArrayGetElement(psa, &ix, &mut vt as *mut _ as *mut c_void);
                        println!("loop {} - hr = 0x{:x}", ix, hr);
                        let vr = Variant::from_c_variant(*vt);
                        vc.push(vr);
                    }
                    RSafeArray::Variants(vc)
                },
                VT_UNKNOWN => unsafe {
                    let mut vc = Vec::new();
                    for ix in lower_bound..upper_bound {
                        let mut pu: *mut IUnknown = ptr::null_mut();
                        let hr = SafeArrayGetElement(psa, &ix, &mut pu as *mut _ as *mut c_void);
                        println!("loop {} - hr = 0x{:x}", ix, hr);
                        vc.push(pu);
                    }
                    RSafeArray::Unknowns(vc, None)
                },
                VT_I1 => unsafe {
                    let mut vc = Vec::new();
                    for ix in lower_bound..upper_bound {
                        let mut i = 0i8;
                        let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void);
                        println!("loop {} - hr = 0x{:x}", ix, hr);
                        vc.push(i);
                    }
                    RSafeArray::Chars(vc)
                },
                VT_UI1 => unsafe {
                    let mut vc = Vec::new();
                    for ix in lower_bound..upper_bound {
                        let mut i = 0u8;
                        let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void);
                        println!("loop {} - hr = 0x{:x}", ix, hr);
                        vc.push(i);
                    }
                    RSafeArray::UChars(vc)
                },
                VT_UI2 => unsafe {
                    let mut vc = Vec::new();
                    for ix in lower_bound..upper_bound {
                        let mut i = 0u16;
                        let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void);
                        println!("loop {} - hr = 0x{:x}", ix, hr);
                        vc.push(i);
                    }
                    RSafeArray::UShorts(vc)
                },
                VT_UI4 => unsafe {
                    let mut vc = Vec::new();
                    for ix in lower_bound..upper_bound {
                        let mut i = 0u32;
                        let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void);
                        println!("loop {} - hr = 0x{:x}", ix, hr);
                        vc.push(i);
                    }
                    RSafeArray::ULongs(vc)
                },
                VT_INT => unsafe {
                    let mut vc = Vec::new();
                    for ix in lower_bound..upper_bound {
                        let mut i = 0i32;

                        let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void);
                        println!("loop {} - hr = 0x{:x}", ix, hr);
                        vc.push(Int(i));
                    }
                    RSafeArray::Ints(vc)
                },
                VT_UINT => unsafe {
                    let mut vc = Vec::new();
                    for ix in lower_bound..upper_bound {
                        let mut i = 0u32;

                        let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void);
                        println!("loop {} - hr = 0x{:x}", ix, hr);
                        vc.push(UInt(i));
                    }
                    RSafeArray::UInts(vc)
                },
                _ => panic!("Not supported!")
            }
        }
        else {
            panic!("Nested safearrays not yet supported.");
        }
    }
}

macro_rules! FROM_RUST_MATCH {
    ($match_name:ident, $psa:ident, {$($enum_name:ident,)*}) => {
        match $match_name {
            $(
                RSafeArray::$enum_name(array) => {
                    for (ix, mut elem) in array.into_iter().enumerate() {
                        let _hr = unsafe {
                            SafeArrayPutElement($psa, &(ix as i32), &mut elem as *mut _ as *mut c_void)
                        };
                    }
                }
            ),*
            _ => panic!("Unsupported variants")
        }
    };
}

impl<T> From<RSafeArray<T>> for *mut SAFEARRAY
    where T: Debug
{
    //consumes underlying object, is this memory safe? 
    fn from(rsa: RSafeArray<T>) -> *mut SAFEARRAY {
        let c_elements: ULONG = rsa.len() as u32;
        let vartype = rsa.vartype();
        let mut sab = SAFEARRAYBOUND {cElements: c_elements, lLbound: 0i32};
        let psa = unsafe {
            SafeArrayCreate(vartype as u16, 1, &mut sab)
        };

        //Allocate struct to safely destroy SAFEARRAY if there is a panic or crash during this process.
        let mut sad = SafeArrayDestructor::new(psa);
        if let RSafeArray::Bools(array) = rsa {
            for (ix, mut elem) in array.into_iter().enumerate() {
                let mut vb_elem: VARIANT_BOOL = if elem {-1} else {0};
                let _hr = unsafe {
                    SafeArrayPutElement(psa, &(ix as i32), &mut vb_elem as *mut _ as *mut c_void)
                };
            }
        }
        else if let RSafeArray::Variants(array) = rsa {
            for (ix, mut elem) in array.into_iter().enumerate() {
                let mut var_elem = elem.into_c_variant();
                let _hr = unsafe {
                    SafeArrayPutElement(psa, &(ix as i32), &mut var_elem as *mut _ as *mut c_void)
                };
            }
        }
        else if let RSafeArray::BStrings(array) = rsa {
            for(ix, mut elem) in array.into_iter().enumerate() {
                let mut bs: bstring::BString = From::from(elem);
                let _hr = unsafe {
                    SafeArrayPutElement(psa, &(ix as i32), &mut bs.as_sys() as *mut _ as *mut c_void)
                };
            }
        }
        else if let RSafeArray::Dispatchs(array, _) = rsa {
            for (ix, mut elem) in array.into_iter().enumerate() {
                let _hr = unsafe {
                    SafeArrayPutElement(psa, &(ix as i32), &mut elem as *mut _ as *mut c_void)
                };
            }
        }
        else if let RSafeArray::Unknowns(array, _) = rsa {
            for (ix, mut elem) in array.into_iter().enumerate() {
                let _hr = unsafe {
                    SafeArrayPutElement(psa, &(ix as i32), &mut elem as *mut _ as *mut c_void)
                };
            }
        }
        else if let RSafeArray::Currencies(array) = rsa {
            for (ix, mut elem) in array.into_iter().enumerate() {
                let mut cy = CY::from(elem);
                let _hr = unsafe {
                    SafeArrayPutElement(psa, &(ix as i32), &mut cy as *mut _ as *mut c_void)
                };
            }
        }
        else if let RSafeArray::Dates(array) = rsa {
            for (ix, mut elem) in array.into_iter().enumerate() {
                let mut dt = elem.0;
                let _hr = unsafe {
                    SafeArrayPutElement(psa, &(ix as i32), &mut dt as *mut _ as *mut c_void)
                };
            }
        }
        else if let RSafeArray::Decimals(array) = rsa {
            for (ix, mut elem) in array.into_iter().enumerate() {
                let mut cdec = build_c_decimal(elem);
                let _hr = unsafe {
                    SafeArrayPutElement(psa, &(ix as i32), &mut cdec as *mut _ as *mut c_void)
                };
            }
        }
        else {
            FROM_RUST_MATCH!{rsa, psa, 
                {Shorts,   Longs,    Floats,   Doubles, Chars, 
                 UChars, UShorts, ULongs, Ints, UInts,}
            };
        }
        sad.inner = ptr::null_mut(); //ensure struct doesn't destroy the safearray
        psa
    }
}