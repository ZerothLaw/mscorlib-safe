use std::marker::PhantomData;
use std::ptr;

use winapi::ctypes::{c_long, c_void};

use winapi::shared::minwindef::{UINT, ULONG};
use winapi::shared::winerror::HRESULT;
use winapi::shared::wtypes::{VARENUM, VARTYPE,    VT_BOOL, VT_DISPATCH, VT_INT,
                             VT_I1,   VT_I2,      VT_I4,   VT_R4,       VT_R8, 
                             VT_UINT, VT_UNKNOWN, VT_UI1,  VT_UI2,      VT_UI4, 
                             VT_VARIANT};
use winapi::shared::wtypes::{VARIANT_BOOL};

use winapi::um::oaidl::{IDispatch, LPSAFEARRAY, LPSAFEARRAYBOUND, SAFEARRAY, SAFEARRAYBOUND, VARIANT};
use winapi::um::unknwnbase::IUnknown;

use new_variant::Variant;


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

pub enum RSafeArray {
    I16(Vec<i16>), //VT_I2
    I32(Vec<i32>), //VT_I4,
    F32(Vec<f32>), //VT_R4, 
    F64(Vec<f64>), //VT_R8, 
    Dispatch(Vec<*mut IDispatch>), //VT_DISPATCH, 
    Bool(Vec<bool>), //VT_BOOL, need to translate between rust bool and VARIANT_BOOL,
    Variant(Vec<Variant>), //VT_VARIANT,
    Unknown(Vec<*mut IUnknown>), //VT_UNKNOWN, 
    SChar(Vec<i8>), //VT_I1, 
    UChar(Vec<u8>), //VT_UI1, 
    UShort(Vec<u16>), //VT_UI2, 
    ULong(Vec<u32>), //VT_UI4, 
    
    #[cfg(target_arch="x86_64")]
    Int(Vec<i64>), //VT_INT, 

    #[cfg(target_arch="x86_64")]
    UInt(Vec<u64>), //VT_UINT, 
    
    #[cfg(target_arch="x86")]
    Int(Vec<i32>), //VT_INT, 
    
    #[cfg(target_arch="x86")]
    UInt(Vec<u32>), //VT_UINT,
}

macro_rules! LEN_SAFEARRAY {
    ($in_name:ident $($enum_name:ident,)*) => {
        match $in_name {
            $(
                RSafeArray::$enum_name(inner) => inner.len()
            ),*
        }
    };
}

macro_rules! VARTYPE_SAFEARRAY {
    ($match_name:ident $($enum_name:ident, $vartype:ident,)*) => {
        match $match_name {
            $(
                RSafeArray::$enum_name(_) => $vartype
            ),*
        }
    };
}

impl RSafeArray {
    pub fn len(&self) -> usize {
        LEN_SAFEARRAY!{self
            I16, I32, F32, F64, Dispatch, Bool, Variant, Unknown, 
            SChar, UChar, UShort, ULong, Int, UInt,
        }
    }
    pub fn vartype(&self) -> VARENUM {
        VARTYPE_SAFEARRAY!{self 
            I16, VT_I2, 
            I32, VT_I4, 
            F32, VT_R4, 
            F64, VT_R8, 
            Dispatch, VT_DISPATCH, 
            Bool, VT_BOOL, 
            Variant, VT_VARIANT, 
            Unknown, VT_UNKNOWN, 
            SChar, VT_I1, 
            UChar, VT_UI1, 
            UShort, VT_UI2, 
            ULong, VT_UI4, 
            Int, VT_INT, 
            UInt, VT_UINT,
        }
    }
}

macro_rules! FROM_IMPLS {
    ($($(#[$attrs:meta])* {$enum_name:ident, $v_type:ty})*) => {
        $(
            $(#[$attrs])*
            impl From<Vec<$v_type>> for RSafeArray {
                fn from(in_vc: Vec<$v_type>) -> RSafeArray {
                    RSafeArray::$enum_name(in_vc)
                }
            }
        )*
    };
}

FROM_IMPLS!{
    {I16, i16}
    {I32, i32}
    {F32, f32}
    {F64, f64}
    {Dispatch, *mut IDispatch}
    {Bool, bool}
    {Variant, Variant}
    {Unknown, *mut IUnknown}
    {SChar, i8}
    {UChar, u8}
    {UShort, u16}
    {ULong, u32}
    #[cfg(target_arch="x86_64")]
    {Int, i64}
    #[cfg(target_arch="x86_64")]
    {UInt, u64}
    #[cfg(target_arch="x86")]
    {Int, i32}
    #[cfg(target_arch="x86")]
    {UInt, u32}
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

impl From<*mut SAFEARRAY> for RSafeArray {
    fn from(psa: *mut SAFEARRAY) -> RSafeArray {
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
                    RSafeArray::I16(vc)
                },
                VT_I4 => unsafe {
                    let mut vc = Vec::new();
                    for ix in lower_bound..upper_bound {
                        let mut i = 0i32;
                        let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void);
                        println!("loop {} - hr = 0x{:x}", ix, hr);
                        vc.push(i);
                    }
                    RSafeArray::I32(vc)
                },
                VT_R4 => unsafe {
                    let mut vc = Vec::new();
                    for ix in lower_bound..upper_bound {
                        let mut i:f32 = 0.0;
                        let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void);
                        println!("loop {} - hr = 0x{:x}", ix, hr);
                        vc.push(i);
                    }
                    RSafeArray::F32(vc)
                },
                VT_R8 => unsafe {
                    let mut vc = Vec::new();
                    for ix in lower_bound..upper_bound {
                        let mut i:f64 = 0.0;
                        let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void);
                        println!("loop {} - hr = 0x{:x}", ix, hr);
                        vc.push(i);
                    }
                    RSafeArray::F64(vc)
                },
                VT_DISPATCH => unsafe {
                    let mut vc = Vec::new();
                    for ix in lower_bound..upper_bound {
                        let mut p: *mut IDispatch = ptr::null_mut();
                        let hr = SafeArrayGetElement(psa, &ix, &mut p as *mut _ as *mut c_void);
                        println!("loop {} - hr = 0x{:x}", ix, hr);
                        vc.push(p);
                    }
                    RSafeArray::Dispatch(vc)
                },
                VT_BOOL => unsafe {
                    let mut vc = Vec::new();
                    for ix in lower_bound..upper_bound {
                        let mut vb: VARIANT_BOOL = 0;
                        let hr = SafeArrayGetElement(psa, &ix, &mut vb as *mut _ as *mut c_void);
                        println!("loop {} - hr = 0x{:x}", ix, hr);
                        vc.push(vb == -1);
                    }
                    RSafeArray::Bool(vc)
                },
                VT_VARIANT => unsafe {
                    let mut vc = Vec::new();
                    for ix in lower_bound..upper_bound {
                        let mut vt: *mut VARIANT = ptr::null_mut();
                        let hr = SafeArrayGetElement(psa, &ix, &mut vt as *mut _ as *mut c_void);
                        println!("loop {} - hr = 0x{:x}", ix, hr);
                        let vr = Variant::from(*vt);
                        vc.push(vr);
                    }
                    RSafeArray::Variant(vc)
                },
                VT_UNKNOWN => unsafe {
                    let mut vc = Vec::new();
                    for ix in lower_bound..upper_bound {
                        let mut pu: *mut IUnknown = ptr::null_mut();
                        let hr = SafeArrayGetElement(psa, &ix, &mut pu as *mut _ as *mut c_void);
                        println!("loop {} - hr = 0x{:x}", ix, hr);
                        vc.push(pu);
                    }
                    RSafeArray::Unknown(vc)
                },
                VT_I1 => unsafe {
                    let mut vc = Vec::new();
                    for ix in lower_bound..upper_bound {
                        let mut i = 0i8;
                        let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void);
                        println!("loop {} - hr = 0x{:x}", ix, hr);
                        vc.push(i);
                    }
                    RSafeArray::SChar(vc)
                },
                VT_UI1 => unsafe {
                    let mut vc = Vec::new();
                    for ix in lower_bound..upper_bound {
                        let mut i = 0u8;
                        let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void);
                        println!("loop {} - hr = 0x{:x}", ix, hr);
                        vc.push(i);
                    }
                    RSafeArray::UChar(vc)
                },
                VT_UI2 => unsafe {
                    let mut vc = Vec::new();
                    for ix in lower_bound..upper_bound {
                        let mut i = 0u16;
                        let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void);
                        println!("loop {} - hr = 0x{:x}", ix, hr);
                        vc.push(i);
                    }
                    RSafeArray::UShort(vc)
                },
                VT_UI4 => unsafe {
                    let mut vc = Vec::new();
                    for ix in lower_bound..upper_bound {
                        let mut i = 0u32;
                        let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void);
                        println!("loop {} - hr = 0x{:x}", ix, hr);
                        vc.push(i);
                    }
                    RSafeArray::ULong(vc)
                },
                VT_INT => unsafe {
                    let mut vc = Vec::new();
                    for ix in lower_bound..upper_bound {
                        #[cfg(target_arch="x86_64")]
                        let mut i = 0i64;

                        #[cfg(target_arch="x86")]
                        let mut i = 0i32;

                        let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void);
                        println!("loop {} - hr = 0x{:x}", ix, hr);
                        vc.push(i);
                    }
                    RSafeArray::Int(vc)
                },
                VT_UINT => unsafe {
                    let mut vc = Vec::new();
                    for ix in lower_bound..upper_bound {
                        #[cfg(target_arch="x86_64")]
                        let mut i = 0u64;

                        #[cfg(target_arch="x86")]
                        let mut i = 0u32;

                        let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void);
                        println!("loop {} - hr = 0x{:x}", ix, hr);
                        vc.push(i);
                    }
                    RSafeArray::UInt(vc)
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

impl From<RSafeArray> for *mut SAFEARRAY{
    //consumes underlying object, is this memory safe? 
    fn from(rsa: RSafeArray) -> *mut SAFEARRAY {
        let c_elements: ULONG = rsa.len() as u32;
        let vartype = rsa.vartype();
        let mut sab = SAFEARRAYBOUND {cElements: c_elements, lLbound: 0i32};
        let psa = unsafe {
            SafeArrayCreate(vartype as u16, 1, &mut sab)
        };
        let mut sad = SafeArrayDestructor::new(psa);
        if let RSafeArray::Bool(array) = rsa {
            for (ix, mut elem) in array.into_iter().enumerate() {
                let mut vb_elem: VARIANT_BOOL = if elem {-1} else {0};
                let _hr = unsafe {
                    SafeArrayPutElement(psa, &(ix as i32), &mut vb_elem as *mut _ as *mut c_void)
                };
            }
        }
        else if let RSafeArray::Variant(array) = rsa {
            for (ix, mut elem) in array.into_iter().enumerate() {
                let mut var_elem = VARIANT::from(elem);
                let _hr = unsafe {
                    SafeArrayPutElement(psa, &(ix as i32), &mut var_elem as *mut _ as *mut c_void)
                };
            }
        }
        else {
            FROM_RUST_MATCH!{rsa, psa, 
                {I16, I32, F32, F64, Dispatch, Unknown, 
                SChar, UChar, UShort, ULong, Int, UInt,}
            };
        }
        sad.inner = ptr::null_mut(); //ensure struct doesn't destroy the safearray
        psa
    }
}