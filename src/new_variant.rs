use std::ops::Deref;
use std::mem;
use winapi::ctypes::c_void;
use winapi::shared::wtypes::{VARENUM, VARIANT_BOOL, VARTYPE,     VT_ARRAY, 
                             VT_BSTR, VT_BOOL,  
                             VT_BYREF,     VT_DISPATCH, VT_I1,    VT_I2,       
                             VT_I4,        VT_I8,       VT_R4,    VT_R8,     
                             VT_UI1,       VT_UI2,      VT_UI4,   VT_UI8,      
                             VT_UNKNOWN,   VT_VARIANT, };
use winapi::um::unknwnbase::IUnknown;
use winapi::um::oaidl::{IDispatch, SAFEARRAY, VARIANT, VARIANT_n1, __tagVARIANT, VARIANT_n3};

use bstring;

#[derive(Debug, Clone, PartialEq, PartialOrd)]
pub enum Primitive {
    I8(i8), //matches general machine understanding of char, not Rust's
    I16(i16), 
    I32(i32), 
    I64(i64),
    F32(f32), 
    F64(f64), 
    U8(u8), //matches general machine understanding of unsigned char, not Rust's
    U16(u16), 
    U32(u32), 
    U64(u64), 
    Bool(bool), 
    BString(String),
    //VT_ERROR, VT_CY, VT_DATE, VT_BSTR
}

use self::Primitive::*;

impl Primitive {
    fn vartype(&self) -> VARTYPE {
        let vt = match self {
            I8(_) => VT_I1, 
            I16(_) => VT_I2, 
            I32(_) => VT_I4, 
            I64(_) => VT_I8, 
            F32(_) => VT_R4, 
            F64(_) => VT_R8, 
            U8(_) => VT_UI1, 
            U16(_) => VT_UI2, 
            U32(_) => VT_UI4, 
            U64(_) => VT_UI8, 
            Bool(_) => VT_BOOL,
            BString(_) => VT_BSTR,
        };
        vt as u16
    }
}

#[derive(Debug, Clone, PartialEq, PartialOrd)]
pub enum Pointer {
    Unknown(*mut IUnknown),
    Dispatch(*mut IDispatch),
    SafeArray(*mut SAFEARRAY), 
    PI8(*mut i8),
    PI16(*mut i16), 
    PI32(*mut i32), 
    PI64(*mut i64),
    PF32(*mut f32), 
    PF64(*mut f64), 
    PU8(*mut u8),
    PU16(*mut u16), 
    PU32(*mut u32), 
    PU64(*mut u64), 
    PBool(*mut bool), 
    ByRef(*mut c_void),
    PVar(*mut VARIANT),
    //n4 -> __tagBRECORD
}
use self::Pointer::*;

impl Pointer {
    fn vartype(&self) -> VARTYPE {
        let vt = match self {
            Unknown(_) => VT_UNKNOWN, 
            Dispatch(_) => VT_DISPATCH, 
            SafeArray(_) => VT_ARRAY, 
            PVar(_) => VT_VARIANT,
            PU8(_) => VT_BYREF | VT_UI1, 
            PI8(_) => VT_BYREF | VT_I1,
            PI16(_) => VT_BYREF | VT_I2, 
            PI32(_) => VT_BYREF | VT_I4, 
            PI64(_) => VT_BYREF | VT_I8, 
            PF32(_) => VT_BYREF | VT_R4,
            PF64(_) => VT_BYREF | VT_R8,
            PU16(_) => VT_BYREF | VT_UI2,
            PU32(_) => VT_BYREF | VT_UI4,
            PU64(_) => VT_BYREF | VT_UI8,   
            PBool(_) => VT_BYREF | VT_BOOL,
            ByRef(_) => VT_BYREF,
        };
        vt as u16
    }
}

#[derive(Debug, Clone, PartialEq, PartialOrd)]
pub enum Variant {
    VariantPointer(Pointer), 
    VariantPrimitive(Primitive)
}

impl Variant {
    fn vartype(&self) -> VARTYPE {
        match self {
            Variant::VariantPointer(inner) => inner.vartype(), 
            Variant::VariantPrimitive(inner) => inner.vartype()
        }
    }
}

impl From<i8> for Variant {
    fn from(in_value: i8) -> Variant {
        Variant::VariantPrimitive(Primitive::I8(in_value))
    }
}

macro_rules! FROM_IMPLS {
    ($(($rule_type:ident => $in_type:ty, $enum_type:ident))*) => {
        $(
            FROM_IMPLS!{@rule $rule_type => $in_type, $enum_type}
        )*
    };
    (@rule prim => $in_type:ty, $enum_type:ident) => {
        impl From<$in_type> for Variant {
            fn from(in_value: $in_type) -> Variant {
                Variant::VariantPrimitive(Primitive::$enum_type(in_value))
            }
        }
    };
    (@rule comp => $in_type:ty, $enum_type:ident) => {
        impl From<*mut $in_type> for Variant {
            fn from(in_value: *mut $in_type) -> Variant {
                Variant::VariantPointer(Pointer::$enum_type(in_value))
            }
        }
    };
}



FROM_IMPLS!{
    (prim => i16,  I16)
    (prim => i32,  I32)
    (prim => i64,  I64)
    (prim => f32,  F32)
    (prim => f64,  F64)
    (prim => u8,   U8)
    (prim => u16,  U16)
    (prim => u32,  U32)
    (prim => u64,  U64)
    (prim => bool, Bool)
    (prim => String, BString)
    (comp => i8,   PI8)
    (comp => i16,  PI16)
    (comp => i32,  PI32)
    (comp => i64,  PI64)
    (comp => f32,  PF32)
    (comp => f64,  PF64)
    (comp => u8,   PU8)
    (comp => u16,  PU16)
    (comp => u32,  PU32)
    (comp => u64,  PU64)
    (comp => bool, PBool)
    (comp => IUnknown,  Unknown)
    (comp => IDispatch, Dispatch)
    (comp => SAFEARRAY, SafeArray)
    (comp => c_void,    ByRef)
    (comp => VARIANT,   PVar)
}

impl From<Variant> for VARIANT {
    fn from(source: Variant) -> VARIANT {
        let vt = source.vartype();
        let n3 = match source {
            Variant::VariantPointer(inner) => {
                let mut n3: VARIANT_n3 = unsafe { mem::zeroed()};
                match inner {
                    Unknown(ptr) => unsafe {
                        let mut n_ptr = n3.punkVal_mut();
                        *n_ptr = ptr;
                    }, 
                    Dispatch(ptr) => unsafe {
                        let mut n_ptr = n3.pdispVal_mut();
                        *n_ptr = ptr;
                    }, 
                    SafeArray(ptr) => unsafe {
                        let mut n_ptr = n3.parray_mut();
                        *n_ptr = ptr;
                    }, 
                    PU8(ptr) => unsafe {
                        let mut n_ptr = n3.pbVal_mut();
                        *n_ptr = ptr;
                    }, 
                    PI16(ptr) => unsafe {
                        let mut n_ptr = n3.piVal_mut();
                        *n_ptr = ptr;
                    }, 
                    PI32(ptr) => unsafe {
                        let mut n_ptr = n3.plVal_mut();
                        *n_ptr = ptr;
                    }, 
                    PI64(ptr) => unsafe {
                        let mut n_ptr = n3.pllVal_mut();
                        *n_ptr = ptr;
                    }, 
                    PF32(ptr) => unsafe {
                        let mut n_ptr = n3.pfltVal_mut();
                        *n_ptr = ptr;
                    }, 
                    PF64(ptr) => unsafe {
                        let mut n_ptr = n3.pdblVal_mut();
                        *n_ptr = ptr;
                    }, 
                    PI8(ptr) => unsafe {
                        let mut n_ptr = n3.pcVal_mut();
                        *n_ptr = ptr;
                    }, 
                    PU16(ptr) => unsafe {
                        let mut n_ptr = n3.puiVal_mut();
                        *n_ptr = ptr;
                    }, 
                    PU32(ptr) => unsafe {
                        let mut n_ptr = n3.pulVal_mut();
                        *n_ptr = ptr;
                    }, 
                    PU64(ptr) => unsafe {
                        let mut n_ptr = n3.pullVal_mut();
                        *n_ptr = ptr;
                    }, 
                    PBool(ptr) => unsafe {
                        let mut n_ptr = n3.pboolVal_mut();
                        //convert Rust bool to COM VARIANT_BOOL
                        // -1 = true, 0 = false
                        // wtf
                        let b_val = *ptr;
                        let mut vb_val: VARIANT_BOOL = if b_val {-1} else {0};
                        *n_ptr = &mut vb_val;
                    }, 
                    ByRef(ptr) => unsafe {
                        let mut n_ptr = n3.byref_mut();
                        *n_ptr = ptr;
                    }, 
                    PVar(ptr) => unsafe {
                        let mut n_ptr = n3.pvarVal_mut();
                        *n_ptr = ptr;
                    }, 
                };
                n3
            }, 
            Variant::VariantPrimitive(inner) => {
                let mut n3: VARIANT_n3 = unsafe { mem::zeroed()};
                match inner {
                    I8(val) => unsafe {
                        let mut n_ptr = n3.cVal_mut();
                        *n_ptr = val;
                    },
                    I16(val) => unsafe {
                        let mut n_ptr = n3.iVal_mut();
                        *n_ptr = val;
                    },
                    I32(val) => unsafe {
                        let mut n_ptr = n3.lVal_mut();
                        *n_ptr = val;
                    },
                    I64(val) => unsafe {
                        let mut n_ptr = n3.llVal_mut();
                        *n_ptr = val;
                    },
                    F32(val) => unsafe {
                        let mut n_ptr = n3.fltVal_mut();
                        *n_ptr = val;
                    },
                    F64(val) => unsafe {
                        let mut n_ptr = n3.dblVal_mut();
                        *n_ptr = val;
                    },
                    U8(val) => unsafe {
                        let mut n_ptr = n3.bVal_mut();
                        *n_ptr = val;
                    },
                    U16(val) => unsafe {
                        let mut n_ptr = n3.uiVal_mut();
                        *n_ptr = val;
                    },
                    U32(val) => unsafe {
                        let mut n_ptr = n3.ulVal_mut();
                        *n_ptr = val;
                    },
                    U64(val) => unsafe {
                        let mut n_ptr = n3.ullVal_mut();
                        *n_ptr = val;
                    },
                    Bool(val) => unsafe {
                        let mut n_ptr = n3.boolVal_mut();
                        let vb_value: VARIANT_BOOL = if val {-1} else {0};
                        *n_ptr = vb_value;
                    },
                    BString(inner) => unsafe {
                        let mut n_ptr = n3.bstrVal_mut();
                        let bs: bstring::BString = From::from(inner);
                        *n_ptr = bs.as_sys();
                    },
                };
                n3
            }
        };
        let tv = __tagVARIANT { vt: vt, 
                                wReserved1: 0, 
                                wReserved2: 0, 
                                wReserved3: 0,
                                n3: n3 };
        let n1 = unsafe {
            let mut n1: VARIANT_n1 = mem::zeroed();
            {
                let mut n_ptr = n1.n2_mut();
                *n_ptr = tv;
            }
            n1
        };

        VARIANT {n1: n1}
    }
}

const VT_PBYTE: VARENUM = VT_BYREF|VT_UI1;
const VT_PSHORT: VARENUM = VT_BYREF|VT_I2;
const VT_PLONG: VARENUM = VT_BYREF|VT_I4;
const VT_PLONGLONG: VARENUM = VT_BYREF|VT_I8;
const VT_PFLOAT: VARENUM = VT_BYREF|VT_R4;
const VT_PDOUBLE: VARENUM = VT_BYREF|VT_R8;
const VT_PBOOL: VARENUM = VT_BYREF|VT_BOOL;
const VT_PVARIANT: VARENUM = VT_BYREF|VT_VARIANT; 
const VT_PCHAR: VARENUM = VT_BYREF|VT_I1;
const VT_PUSHORT: VARENUM = VT_BYREF|VT_UI2;
const VT_PULONG: VARENUM = VT_BYREF|VT_UI4;
const VT_PULONGLONG: VARENUM = VT_BYREF|VT_UI8;


//VTYPE_MAP!{vt as u32, n3, 
//    (prim, VT_I8, llVal, I64)
//}
macro_rules! VTYPE_MAP {
    ($match_name:expr, $union_name:ident, $(($rule:ident,$var_type:ident, $union_func:ident, $enum_name:ident) )*) => {
        match $match_name {
            $(
                $var_type => unsafe {
                    let val = {
                        let n_ptr = $union_name.$union_func();
                        *n_ptr
                    };
                    VTYPE_MAP!{@rule $rule, $enum_name, val}
                }
            ),*
            _ => panic!("Not supported.")
        }
    };
    (@rule prim, $enum_name:ident, $var_name:ident) => {
        Variant::VariantPrimitive(Primitive::$enum_name($var_name))
    };
    (@rule comp, $enum_name:ident, $var_name:ident) => {
        Variant::VariantPointer(Pointer::$enum_name($var_name))
    };
}

impl From<VARIANT> for Variant {
    fn from(vt: VARIANT) -> Variant {
        let mut n1 = vt.n1;
        let n2_mut = unsafe {
            n1.n2_mut()
        };
        let vt: VARTYPE = n2_mut.vt;
        let mut n3 = n2_mut.n3;
        //have to special case bool/pbool conversions because of VARIANT_BOOL
        if vt as u32 == VT_BOOL {
            let val = unsafe {
                let n_ptr = n3.boolVal();
                *n_ptr
            };
            let b_val = val == -1;
            Variant::VariantPrimitive(Primitive::Bool(b_val))
        }
        else if vt as u32 == VT_PBOOL {
            let val = unsafe {
                let n_ptr = n3.pboolVal();
                **n_ptr
            };
            let mut b_val = val == -1;
            Variant::VariantPointer(Pointer::PBool(&mut b_val))
        }
        else if vt as u32 == VT_BSTR {
            let val = unsafe {
                let n_ptr = n3.bstrVal();
                *n_ptr
            };
            let bs: bstring::BString = bstring::BString::from_ptr_safe(val);
            Variant::VariantPrimitive(Primitive::BString(bs.to_string()))
        }
        else {
            VTYPE_MAP!{vt as u32, n3,
                (prim, VT_I8, llVal, I64)
                (prim, VT_I4, lVal,  I32)
                (prim, VT_UI1, bVal, U8)
                (prim, VT_I2, iVal, I16)
                (prim, VT_R4, fltVal, F32)
                (prim, VT_R8, dblVal, F64)
                //(prim, VT_BOOL, boolVal, Bool)
                (prim, VT_I1, cVal, I8)
                (prim, VT_UI2, uiVal, U16)
                (prim, VT_UI4, ulVal, U32)
                (prim, VT_UI8, ullVal, U64)
                (comp, VT_UNKNOWN, punkVal_mut, Unknown)
                (comp, VT_DISPATCH, pdispVal_mut, Dispatch)
                (comp, VT_ARRAY, parray_mut, SafeArray)
                (comp, VT_PBYTE, pbVal, PU8)
                (comp, VT_PSHORT, piVal, PI16)
                (comp, VT_PLONG, plVal, PI32)
                (comp, VT_PLONGLONG, pllVal, PI64)
                (comp, VT_PFLOAT, pfltVal, PF32)
                (comp, VT_PDOUBLE, pdblVal, PF64)
                //(comp, VT_PBOOL, pboolVal, PBool)
                (comp, VT_PVARIANT, pvarVal, PVar)
                (comp, VT_PCHAR, pcVal, PI8)
                (comp, VT_PUSHORT, puiVal, PU16)
                (comp, VT_PULONG, pulVal, PU32)
                (comp, VT_PULONGLONG, pullVal, PU64)
            }
        }
    }
}