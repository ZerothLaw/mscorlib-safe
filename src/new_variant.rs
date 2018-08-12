use std::mem;

use rust_decimal::Decimal;

use winapi::ctypes::c_void;
use winapi::shared::wtypes::{BSTR,     CY,     DATE, DECIMAL,     DECIMAL_NEG,
                             VARENUM,  VARIANT_BOOL, VARTYPE,     VT_ARRAY,
                             VT_BSTR,  VT_BOOL,      VT_BYREF,    VT_CY, 
                             VT_DATE,  VT_DECIMAL,   VT_DISPATCH, VT_EMPTY, 
                             VT_ERROR, VT_INT,      VT_I1,        VT_I2,       
                             VT_I4,    VT_I8,       VT_NULL,      VT_R4,    
                             VT_R8,    VT_UINT,     VT_UI1,       VT_UI2,      
                             VT_UI4,   VT_UI8,      VT_UNKNOWN,   VT_VARIANT, };
use winapi::shared::wtypesbase::SCODE;
use winapi::um::unknwnbase::IUnknown;
use winapi::um::oaidl::{IDispatch, SAFEARRAY, VARIANT, VARIANT_n1, __tagVARIANT, VARIANT_n3};

use bstring;
use new_safearray::{LPSAFEARRAY, RSafeArray};

const VT_PBYTE: VARENUM = VT_BYREF|VT_UI1;
const VT_PSHORT: VARENUM = VT_BYREF|VT_I2;
const VT_PLONG: VARENUM = VT_BYREF|VT_I4;
const VT_PLONGLONG: VARENUM = VT_BYREF|VT_I8;
const VT_PFLOAT: VARENUM = VT_BYREF|VT_R4;
const VT_PDOUBLE: VARENUM = VT_BYREF|VT_R8;
const VT_PBOOL: VARENUM = VT_BYREF|VT_BOOL;
const VT_PERROR: VARENUM = VT_BYREF|VT_ERROR;
const VT_PCY: VARENUM = VT_BYREF|VT_CY;
const VT_PDATE: VARENUM = VT_BYREF|VT_DATE;
const VT_PBSTR: VARENUM = VT_BYREF|VT_BSTR;
const VT_PUNKNOWN: VARENUM = VT_BYREF|VT_UNKNOWN;
const VT_PDISPATCH: VARENUM = VT_BYREF|VT_DISPATCH;
const VT_PARRAY: VARENUM = VT_BYREF|VT_ARRAY;
const VT_PVARIANT: VARENUM = VT_BYREF|VT_VARIANT; 
const VT_PDECIMAL: VARENUM = VT_BYREF|VT_DECIMAL; 
const VT_PCHAR: VARENUM = VT_BYREF|VT_I1;
const VT_PUSHORT: VARENUM = VT_BYREF|VT_UI2;
const VT_PULONG: VARENUM = VT_BYREF|VT_UI4;
const VT_PULONGLONG: VARENUM = VT_BYREF|VT_UI8;
const VT_PINT: VARENUM = VT_BYREF|VT_INT; 
const VT_PUINT: VARENUM = VT_BYREF|VT_UINT;

#[derive(Debug, Clone, PartialEq, PartialOrd)]
pub struct Currency(i64);

impl From<i64> for Currency {
    fn from(source: i64) -> Currency {
        Currency(source)
    }
}

impl From<Currency> for CY {
    fn from(source: Currency) -> CY {
        CY {int64: source.0}
    }
}

impl From<CY> for Currency {
    fn from(cy: CY) -> Currency {
        Currency(cy.int64)
    }
}

#[derive(Debug, Clone, PartialEq, PartialOrd)]
pub struct SCode(i32);

#[derive(Debug, Clone, PartialEq, PartialOrd)]
pub struct Int(i32);

impl From<i32> for Int {
    fn from(i: i32) -> Int {
        Int(i)
    }
}

#[derive(Debug, Clone, PartialEq, PartialOrd)]
pub struct UInt(u32);

impl From<u32> for UInt {
    fn from(u: u32) -> UInt {
        UInt(u)
    }
}

#[derive(Debug, Clone, PartialEq, PartialOrd)]
pub struct Date(DATE);

fn build_c_decimal(dec: Decimal) -> DECIMAL {
    let scale = dec.scale() as u8;
    let sign = if dec.is_sign_positive() {0} else {DECIMAL_NEG};
    let serial = dec.serialize();
    let lo: u64 = (serial[4]  as u64)        + 
                 ((serial[5]  as u64) << 8)  + 
                 ((serial[6]  as u64) << 16) + 
                 ((serial[7]  as u64) << 24) + 
                 ((serial[8]  as u64) << 32) +
                 ((serial[9]  as u64) << 40) +
                 ((serial[10] as u64) << 48) + 
                 ((serial[11] as u64) << 56);
    let hi: u32 = (serial[12] as u32)        +
                 ((serial[13] as u32) << 8)  +
                 ((serial[14] as u32) << 16) +
                 ((serial[15] as u32) << 24);
    DECIMAL {
        wReserved: 0, 
        scale: scale, 
        sign: sign, 
        Hi32: hi, 
        Lo64: lo
    }
}

fn build_rust_decimal(dec: DECIMAL) -> Decimal {
    let sign = if dec.sign == DECIMAL_NEG {true} else {false};
    Decimal::from_parts((dec.Lo64 & 0xFFFFFFFF) as u32, 
                        ((dec.Lo64 >> 32) & 0xFFFFFFFF) as u32, 
                        dec.Hi32, 
                        sign,
                        dec.scale as u32 ) 
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn c_decimal() {
        let d = Decimal::new(0xFFFFFFFFFFFF, 0);
        let d = d * Decimal::new(0xFFFFFFFF, 0);
        assert_eq!(d.is_sign_positive(), true);
        assert_eq!(format!("{}", d), "1208925819333149903028225" );
        
        let c = build_c_decimal(d);
        println!("({}, {}, {}, {})", c.Hi32, c.Lo64, c.scale, c.sign);
        println!("{:?}", d.serialize());
        assert_eq!(c.Hi32, 65535);
        assert_eq!(c.Lo64, 18446462594437873665); 
        assert_eq!(c.scale, 0);
        assert_eq!(c.sign, 0);
    }

    #[test]
    fn rust_decimal_from() {
        let d = DECIMAL {
            wReserved: 0, 
            scale: 0, 
            sign: 0, 
            Hi32: 65535, 
            Lo64: 18446462594437873665
        };
        let new_d = build_rust_decimal(d);
        println!("{:?}", new_d.serialize());
       // assert_eq!(new_d.is_sign_positive(), true);
        assert_eq!(format!("{}", new_d), "1208925819333149903028225"  );
    }
}

#[derive(Debug, Clone, PartialEq, PartialOrd)]
pub enum Variant {
    LongLong(i64), 
    Long(i32), 
    Byte(u8), 
    Short(i16), 
    Float(f32), 
    Double(f64), 
    Bool(bool), 
    ErrorCode(SCode), 
    Currency(Currency), 
    Date(Date), 
    BString(String),
    Unknown(*mut IUnknown), 
    Dispatch(*mut IDispatch), 
    Array(RSafeArray), 
    //BRecord needs to be handled eventually
    PByte(Box<u8>), 
    PShort(Box<i16>), 
    PLong(Box<i32>), 
    PLongLong(Box<i64>), 
    PFloat(Box<f32>),
    PDouble(Box<f64>), 
    PBool(Box<bool>), 
    PErrorCode(Box<SCode>), 
    PCurrency(Box<Currency>), 
    PDate(Box<Date>), 
    PBString(Box<String>), 
    PUnknown(Box<*mut IUnknown>), 
    PDispatch(Box<*mut IDispatch>), 
    PArray(Box<RSafeArray>), 
    PVariant(Box<Variant>), 
    ByRef(*mut c_void), 
    Char(i8), 
    UShort(u16), 
    ULong(u32), 
    ULongLong(u64), 
    Int(Int), 
    UInt(UInt), 
    PDecimal(Box<Decimal>), 
    PChar(Box<i8>), 
    PUShort(Box<u16>), 
    PULong(Box<u32>), 
    PULongLong(Box<u64>), 
    PInt(Box<Int>), 
    PUInt(Box<UInt>),
    //non n3 variants 
    Decimal(Decimal),
    Empty(()), 
    Null(())
}

/*
match vt as u32 {
    VT_I8 => BRANCH_FROM_RAW!{S, LongLong(n3, llVal)},
    VT_BOOL => BRANCH_FROM_RAW!{C, |val| {Variant::Bool(val == -1)};(boolVal)}, 
}
*/

macro_rules! BRANCH_FROM_RAW {
    (S, $enum_name:ident($union_name:ident, $union_func:ident)) => {
        {
            let val = unsafe {
                let n_ptr = $union_name.$union_func();
                *n_ptr 
            };
            Variant::$enum_name(val)
        }
    };
    (C, $closure:expr; ($union_name:ident, $union_func:ident) ) => {
        {
            let val = unsafe {
                let n_ptr = $union_name.$union_func();
                *n_ptr 
            };
            $closure(val)
        }
    };
}

impl Variant {
    pub fn vartype(&self) -> VARTYPE {
        let vt = match self {
            Variant::LongLong(_) => VT_I8, 
            Variant::Long(_) => VT_I4, 
            Variant::Byte(_) => VT_UI1, 
            Variant::Short(_) => VT_I2, 
            Variant::Float(_) => VT_R4, 
            Variant::Double(_) => VT_R8, 
            Variant::Bool(_) => VT_BOOL, 
            Variant::ErrorCode(_) => VT_ERROR, 
            Variant::Currency(_) => VT_CY, 
            Variant::Date(_) => VT_DATE, 
            Variant::BString(_) => VT_BSTR, 
            Variant::Unknown(_) => VT_UNKNOWN, 
            Variant::Dispatch(_) => VT_DISPATCH,
            Variant::Array(_) => VT_ARRAY, 
            Variant::PByte(_) => VT_PBYTE, 
            Variant::PShort(_) => VT_PSHORT, 
            Variant::PLong(_) => VT_PLONG, 
            Variant::PLongLong(_) => VT_PLONGLONG, 
            Variant::PFloat(_) => VT_PFLOAT, 
            Variant::PDouble(_) => VT_PDOUBLE, 
            Variant::PBool(_) => VT_PBOOL, 
            Variant::PErrorCode(_) => VT_PERROR, 
            Variant::PCurrency(_) => VT_PCY, 
            Variant::PDate(_) => VT_PDATE, 
            Variant::PBString(_) => VT_PBSTR,
            Variant::PUnknown(_) => VT_PUNKNOWN, 
            Variant::PDispatch(_) => VT_PDISPATCH, 
            Variant::PArray(_) => VT_PARRAY, 
            Variant::PVariant(_) => VT_PVARIANT,
            Variant::ByRef(_) => VT_BYREF, 
            Variant::Char(_) => VT_I1, 
            Variant::UShort(_) => VT_UI2, 
            Variant::ULong(_) => VT_UI4, 
            Variant::ULongLong(_) => VT_UI8, 
            Variant::Int(_) => VT_INT, 
            Variant::UInt(_) => VT_UINT, 
            Variant::PDecimal(_) => VT_PDECIMAL, 
            Variant::PChar(_) => VT_PCHAR, 
            Variant::PUShort(_) => VT_PUSHORT,
            Variant::PULong(_) => VT_PULONG, 
            Variant::PULongLong(_) => VT_PULONGLONG,
            Variant::PInt(_) => VT_PINT, 
            Variant::PUInt(_) => VT_PUINT,
            Variant::Decimal(_) => VT_DECIMAL,
            Variant::Empty(_) => VT_EMPTY,
            Variant::Null(_) => VT_NULL
        };
        vt as u16
    }

    pub fn from_c_variant(vt: VARIANT) -> Variant {
        let mut n1 = vt.n1;
        
        let vt: VARTYPE = unsafe {
            let n2_mut = n1.n2_mut();
            n2_mut.vt
        };
        let mut n3 = unsafe {
            let n2_mut = n1.n2_mut();   
            n2_mut.n3
        };
        //have to special case bool/pbool conversions because of VARIANT_BOOL
        match vt as u32 {
            VT_I8 => BRANCH_FROM_RAW!{S, LongLong(n3, llVal)},
            VT_I4 => BRANCH_FROM_RAW!{S, Long(n3, lVal)}, 
            VT_UI1 => BRANCH_FROM_RAW!{S, Byte(n3, bVal)},
            VT_I2 => BRANCH_FROM_RAW!{S, Short(n3, iVal)},
            VT_R4 => BRANCH_FROM_RAW!{S, Float(n3, fltVal)}, 
            VT_R8 => BRANCH_FROM_RAW!{S, Double(n3, dblVal)}, 
            VT_BOOL => BRANCH_FROM_RAW!{C, |val| {
                Variant::Bool(val == -1)
            };(n3, boolVal)}, 
            VT_ERROR => BRANCH_FROM_RAW!{C, |val| {
                Variant::ErrorCode(SCode(val))
            }; (n3, scode)},
            VT_CY => BRANCH_FROM_RAW!{C, |val| {
                Variant::Currency(Currency::from(val))
            }; (n3, cyVal)},
            VT_DATE => BRANCH_FROM_RAW!{C, |val| {
                Variant::Date(Date(val))
            }; (n3, date)},
            VT_BSTR => BRANCH_FROM_RAW!{C, |val| {
                let bs = bstring::BString::from_ptr_safe(val);
                Variant::BString(bs.to_string())
            }; (n3, bstrVal)},
            VT_UNKNOWN => BRANCH_FROM_RAW!{S, Unknown(n3, punkVal)}, 
            VT_DISPATCH => BRANCH_FROM_RAW!{S, Dispatch(n3, pdispVal)}, 
            VT_ARRAY => BRANCH_FROM_RAW!{C, |val| {
                let rsa: RSafeArray<i32> = RSafeArray::from(val);
                Variant::Array(rsa)
            }; (n3, parray)}, 
            VT_PBYTE => BRANCH_FROM_RAW!{C, |val: *mut u8| {
                Variant::PByte(Box::new(unsafe{*val}))
            }; (n3, pbVal) },
            VT_PSHORT => BRANCH_FROM_RAW!{C, |val: *mut i16| {
                Variant::PShort(Box::new(unsafe{*val}))
            }; (n3, piVal) },
            VT_PLONG => BRANCH_FROM_RAW!{C, |val: *mut i32| {
                Variant::PLong(Box::new(unsafe{*val}))
            } ; (n3, plVal) },
            VT_PLONGLONG => BRANCH_FROM_RAW!{C, |val: *mut i64| {
                Variant::PLongLong(Box::new(unsafe{*val}))
            }; (n3, pllVal)}, 
            VT_PFLOAT => BRANCH_FROM_RAW!{C, |val: *mut f32| {
                Variant::PFloat(Box::new(unsafe{*val}))
            }; (n3, pfltVal)},
            VT_PDOUBLE => BRANCH_FROM_RAW!{C, |val: *mut f64| {
                Variant::PDouble(Box::new(unsafe{*val}))
            }; (n3, pdblVal)},
            VT_PBOOL => BRANCH_FROM_RAW!{C, |val: *mut VARIANT_BOOL| {
                Variant::PBool(Box::new(unsafe{*val}  == -1))
            }; (n3, pboolVal)},
            VT_PERROR => BRANCH_FROM_RAW!{C, |val: *mut SCODE| {
                Variant::PErrorCode(Box::new(SCode(unsafe{*val})))
            } ; (n3, pscode)},
            VT_PCY => BRANCH_FROM_RAW!{C, |val: *mut CY| {
                Variant::PCurrency(Box::new(Currency::from(unsafe{*val})))
            } ; (n3, pcyVal)},
            VT_PDATE => BRANCH_FROM_RAW!{C, |val: *mut DATE| {
                Variant::PDate(Box::new(Date(unsafe{*val})))
            }; (n3, pdate)},
            VT_PBSTR => BRANCH_FROM_RAW!{C, |val: *mut BSTR| {
                Variant::PBString(
                    Box::new(
                        bstring::BString::from_ptr_safe(
                            unsafe{
                                *val
                            }
                        ).to_string()
                    )
                )
            }; (n3, pbstrVal )},
            VT_PUNKNOWN => BRANCH_FROM_RAW!{C, |val: *mut *mut IUnknown| {
                Variant::PUnknown(Box::new(unsafe{*val}))
                }; (n3, ppunkVal_mut)},
            VT_PDISPATCH => BRANCH_FROM_RAW!{C, |val: *mut *mut IDispatch| {
                Variant::PDispatch(Box::new(unsafe{*val}))
                }; (n3, ppdispVal_mut)},
            VT_PARRAY => BRANCH_FROM_RAW!{C, |val: *mut *mut SAFEARRAY| {
                Variant::PArray(Box::new(RSafeArray::from(unsafe{*val})))
                }; (n3, pparray)},
            VT_PVARIANT => BRANCH_FROM_RAW!{C, |val: *mut VARIANT | {
                Variant::PVariant(Box::new(Variant::from_c_variant(unsafe{*val})))
                }; (n3, pvarVal)},
            VT_BYREF => BRANCH_FROM_RAW!{C, |val: *mut c_void| {
                Variant::ByRef(val)}; (n3, byref)
                },
            VT_I1 => BRANCH_FROM_RAW!{S, Char(n3, cVal)},
            VT_UI2 => BRANCH_FROM_RAW!{S, UShort(n3, uiVal)},
            VT_UI4 => BRANCH_FROM_RAW!{S, ULong(n3, ulVal)},
            VT_UI8 => BRANCH_FROM_RAW!{S, ULongLong(n3, ullVal)},
            VT_INT => BRANCH_FROM_RAW!{C, |val| { 
                Variant::Int(Int(val)) 
                };(n3, intVal)},
            VT_UINT => BRANCH_FROM_RAW!{C, |val|{
                Variant::UInt(UInt(val)) 
                };(n3, uintVal)},
            VT_PDECIMAL => BRANCH_FROM_RAW!{C, |val: *mut DECIMAL| {
                Variant::PDecimal(Box::new(build_rust_decimal(unsafe{*val})))
                }; (n3, pdecVal)},
            VT_PCHAR => BRANCH_FROM_RAW!{C, |val: *mut i8| {
                Variant::PChar(Box::new(unsafe{*val}))
                }; (n3, pcVal)},
            VT_PUSHORT => BRANCH_FROM_RAW!{C, |val: *mut u16| {
                Variant::PUShort(Box::new(unsafe{*val}))
                }; (n3, puiVal)},
            VT_PULONG => BRANCH_FROM_RAW!{C, |val: *mut u32| {
                Variant::PULong(Box::new(unsafe{*val}))
                }; (n3, pulVal)},
            VT_PULONGLONG => BRANCH_FROM_RAW!{C, |val: *mut u64| {
                Variant::PULongLong(Box::new(unsafe{*val}))
                }; (n3, pullVal)},
            VT_PINT => BRANCH_FROM_RAW!{C, |val: *mut i32| {
                Variant::PInt(Box::new(Int(unsafe{*val})))
                }; (n3, pintVal)},
            VT_PUINT => BRANCH_FROM_RAW!{C, |val: *mut u32| {
                Variant::PUInt(Box::new(UInt(unsafe{*val})))
                }; (n3, puintVal)},
            VT_DECIMAL => BRANCH_FROM_RAW!{C, |val: DECIMAL| {
                let new_val = build_rust_decimal(val);
                Variant::Decimal(new_val)
            }; (n1, decVal)},
            VT_EMPTY => {
                Variant::Empty(())
            }, 
            VT_NULL => {
                Variant::Null(())
            },
            _ => panic!("Unknown vartype: {}", vt)
        }
    }

    pub fn into_c_variant(self) -> VARIANT {
        let vt = self.vartype();
        let mut n3: VARIANT_n3 = unsafe { mem::zeroed()};
        let mut n1: VARIANT_n1 = unsafe { mem::zeroed()};
        
        match self {
            Variant::LongLong(val) => unsafe {
                let mut n_ptr = n3.llVal_mut();
                *n_ptr = val;
            },
            Variant::Long(val) => unsafe {
                let mut n_ptr = n3.lVal_mut();
                *n_ptr = val;
            },
            Variant::Byte(val) => unsafe {
                let mut n_ptr = n3.bVal_mut();
                *n_ptr = val;
            },
            Variant::Short(val) => unsafe {
                let mut n_ptr = n3.iVal_mut();
                *n_ptr = val;
            },
            Variant::Float(val) => unsafe {
                let mut n_ptr = n3.fltVal_mut();
                *n_ptr = val;
            },
            Variant::Double(val) => unsafe {
                let mut n_ptr = n3.dblVal_mut();
                *n_ptr = val;
            },
            Variant::Bool(val) => unsafe {
                let mut n_ptr = n3.boolVal_mut();
                let vb_value: VARIANT_BOOL = if val {-1} else {0};
                *n_ptr = vb_value;
            },
            Variant::ErrorCode(SCode(val)) => unsafe {
                let mut n_ptr = n3.scode_mut();
                let sc_val: SCODE = val;
                *n_ptr = sc_val;
            },
            Variant::Currency(val) => unsafe {
                let mut n_ptr = n3.cyVal_mut();
                *n_ptr = CY::from(val);
            },
            Variant::Date(Date(val)) => unsafe {
                let mut n_ptr = n3.date_mut();
                let dt_val: DATE = val;
                *n_ptr = dt_val;
            },
            Variant::BString(inner) => unsafe {
                let mut n_ptr = n3.bstrVal_mut();
                let bs: bstring::BString = From::from(inner);
                *n_ptr = bs.as_sys();
            },
            Variant::Unknown(ptr) => unsafe {
                let mut n_ptr = n3.punkVal_mut();
                *n_ptr = ptr;
            }, 
            Variant::Dispatch(ptr) => unsafe {
                let mut n_ptr = n3.pdispVal_mut();
                *n_ptr = ptr;
            }, 
            Variant::Array(array) => unsafe {
                let mut n_ptr = n3.parray_mut();
                *n_ptr = LPSAFEARRAY::from(array);
            }, 
            Variant::PByte(boxed_ptr) => unsafe {
                let mut n_ptr = n3.pbVal_mut();
                *n_ptr = Box::into_raw(boxed_ptr);
            }, 
            Variant::PShort(boxed_ptr) => unsafe {
                let mut n_ptr = n3.piVal_mut();
                *n_ptr = Box::into_raw(boxed_ptr);
            }, 
            Variant::PLong(boxed_ptr) => unsafe {
                let mut n_ptr = n3.plVal_mut();
                *n_ptr = Box::into_raw(boxed_ptr);
            }, 
            Variant::PLongLong(boxed_ptr) => unsafe {
                let mut n_ptr = n3.pllVal_mut();
                *n_ptr = Box::into_raw(boxed_ptr);
            }
            Variant::PFloat(boxed_ptr) => unsafe {
                let mut n_ptr = n3.pfltVal_mut();
                *n_ptr = Box::into_raw(boxed_ptr);
            }, 
            Variant::PDouble(boxed_ptr) => unsafe {
                let mut n_ptr = n3.pdblVal_mut();
                *n_ptr = Box::into_raw(boxed_ptr);
            }, 
            Variant::PBool(boxed_ptr) => unsafe {
                let mut n_ptr = n3.pboolVal_mut();
                //convert Rust bool to COM VARIANT_BOOL
                // -1 = true, 0 = false
                // wtf
                let b_val = *boxed_ptr;
                let mut vb_val: VARIANT_BOOL = if b_val {-1} else {0};
                *n_ptr = &mut vb_val;
            }, 
            Variant::PErrorCode(mut boxed_ptr) => unsafe {
                let mut n_ptr = n3.pscode_mut();
                *n_ptr = &mut (*boxed_ptr).0 as *mut SCODE; 
            },
            Variant::PCurrency(boxed_ptr) => unsafe {
                let mut n_ptr = n3.pcyVal_mut();
                let cy = *boxed_ptr;
                *n_ptr = &mut CY::from(cy) as *mut CY;
            },
            Variant::PDate(boxed_ptr) => unsafe {
                let mut n_ptr = n3.pdate_mut();
                let mut dt = (*boxed_ptr).0;
                *n_ptr = &mut dt as *mut DATE;
            },
            Variant::PBString(boxed_ptr) => unsafe {
                let mut n_ptr = n3.pbstrVal_mut();
                let mut bs: bstring::BString = From::from(*boxed_ptr);
                *n_ptr = &mut bs.as_sys() as *mut BSTR;
            }, 
            Variant::PUnknown(boxed_ptr) => unsafe {
                let mut n_ptr = n3.ppunkVal_mut();
                *n_ptr = Box::into_raw(boxed_ptr);
            }, 
            Variant::PDispatch(boxed_ptr) => unsafe {
                let mut n_ptr = n3.ppdispVal_mut();
                *n_ptr = Box::into_raw(boxed_ptr);
            }, 
            Variant::PArray(boxed_ptr) => unsafe {
                let mut n_ptr = n3.pparray_mut();
                let mut psa: *mut SAFEARRAY = <*mut SAFEARRAY>::from(*boxed_ptr);
                *n_ptr = &mut psa;
            },
            Variant::PVariant(boxed_ptr) => unsafe {
                let mut n_ptr = n3.pvarVal_mut();
                let mut vt = Variant::into_c_variant(*boxed_ptr);
                *n_ptr = &mut vt;
            }
            Variant::ByRef(ptr) => unsafe {
                let mut n_ptr = n3.byref_mut();
                *n_ptr = ptr;
            },
            Variant::Char(val) => unsafe {
                let mut n_ptr = n3.cVal_mut();
                *n_ptr = val;
            },
            Variant::UShort(val) => unsafe {
                let mut n_ptr = n3.uiVal_mut();
                *n_ptr = val;
            },
            Variant::ULong(val) => unsafe {
                let mut n_ptr = n3.ulVal_mut();
                *n_ptr = val;
            },
            Variant::ULongLong(val) => unsafe {
                let mut n_ptr = n3.ullVal_mut();
                *n_ptr = val;
            },
            Variant::Int(Int(ival)) => unsafe {
                let mut n_ptr = n3.intVal_mut();
                *n_ptr = ival;
            }
            Variant::UInt(UInt(uval)) => unsafe {
                let mut n_ptr = n3.uintVal_mut();
                *n_ptr = uval;
            },
            Variant::PDecimal(boxed_dec) => unsafe {
                let mut n_ptr = n3.pdecVal_mut();
                let bd = *boxed_dec;
                let mut d = build_c_decimal(bd);
                let b = Box::new(d);
                *n_ptr = Box::into_raw(b);
            },
            Variant::PChar(boxed_ptr) => unsafe {
                let mut n_ptr = n3.pcVal_mut();
                *n_ptr = Box::into_raw(boxed_ptr);
            }, 
            Variant::PUShort(boxed_ptr) => unsafe {
                let mut n_ptr = n3.puiVal_mut();
                *n_ptr = Box::into_raw(boxed_ptr);
            }, 
            Variant::PULong(boxed_ptr) => unsafe {
                let mut n_ptr = n3.pulVal_mut();
                *n_ptr = Box::into_raw(boxed_ptr);
            }, 
            Variant::PULongLong(boxed_ptr) => unsafe {
                let mut n_ptr = n3.pullVal_mut();
                *n_ptr = Box::into_raw(boxed_ptr);
            }, 
            Variant::PInt(mut boxed_ptr) => unsafe {
                let mut n_ptr = n3.pintVal_mut();
                *n_ptr = &mut (*boxed_ptr).0;
            },
            Variant::PUInt(mut boxed_ptr) => unsafe {
                let mut n_ptr = n3.puintVal_mut();
                *n_ptr = &mut (*boxed_ptr).0;
            },
            Variant::Decimal(dec) => unsafe {
                let mut n_ptr = n1.decVal_mut();
                *n_ptr = build_c_decimal(dec);
            },
            Variant::Empty(_) => {}, 
            Variant::Null(_) => {},
        };

        let tv = __tagVARIANT { vt: vt, 
                                wReserved1: 0, 
                                wReserved2: 0, 
                                wReserved3: 0,
                                n3: n3 };
        unsafe {
            let n_ptr = n1.n2_mut();
            *n_ptr = tv;
        };

        VARIANT {n1: n1}
    }
}

impl From<i32> for SCode {
    fn from(source: i32) -> SCode {
        SCode(source)
    }
}

macro_rules! FROM_IMPLS {
    ($( $(#[$attrs:meta])* ($in_type:ty, $enum_type:ident)),*) => {
        $(
            $(
                #[$attrs]
            )*
            impl From<$in_type> for Variant {
                fn from(in_value: $in_type) -> Variant {
                    Variant::$enum_type(in_value)
                }
            }
        )*
    };
}

FROM_IMPLS!{
    (i64,  LongLong),
    (i32, Long),
    (u8, Byte),
    (i16, Short),
    (f32, Float),
    (f64, Double),
    (bool, Bool),
    (SCode, ErrorCode), 
    (Currency, Currency),
    (Date, Date),
    (String, BString),
    (*mut IUnknown, Unknown),
    (*mut IDispatch, Dispatch),
    (RSafeArray, Array),
    (Box<u8>, PByte),
    (Box<i16>, PShort),
    (Box<i32>, PLong),
    (Box<i64>, PLongLong),
    (Box<f32>, PFloat), 
    (Box<f64>, PDouble),
    (Box<bool>, PBool), 
    (Box<SCode>, PErrorCode),
    (Box<Currency>, PCurrency), 
    (Box<Date>, PDate), 
    (Box<String>, PBString), 
    (Box<*mut IUnknown>, PUnknown), 
    (Box<*mut IDispatch>, PDispatch), 
    (Box<RSafeArray>, PArray), 
    (Box<Variant>, PVariant), 
    (*mut c_void, ByRef), 
    (i8, Char), 
    (u16, UShort), 
    (u32, ULong),
    (u64, ULongLong), 
    (Int, Int), 
    (UInt, UInt), 
    (Box<Decimal>, PDecimal), 
    (Box<i8>, PChar), 
    (Box<u16>, PUShort), 
    (Box<u32>, PULong),
    (Box<u64>, PULongLong),
    (Box<Int>, PInt), 
    (Box<UInt>, PUInt),
    (Decimal, Decimal)
}

impl From<Box<IUnknown>> for Variant {
    fn from(boxed_ptr: Box<IUnknown>) -> Variant {
        Variant::Unknown(Box::into_raw(boxed_ptr))
    }
}

impl From<Box<IDispatch>> for Variant {
    fn from(boxed_ptr: Box<IDispatch>) -> Variant {
        Variant::Dispatch(Box::into_raw(boxed_ptr))
    }
}

impl From<Box<Box<IUnknown>>> for Variant {
    fn from(boxed_ptr: Box<Box<IUnknown>>) -> Variant {
        Variant::PUnknown(Box::new(Box::into_raw(*boxed_ptr)))
    }
}

impl From<Box<Box<IDispatch>>> for Variant {
    fn from(boxed_ptr: Box<Box<IDispatch>>) -> Variant {
        Variant::PDispatch(Box::new(Box::into_raw(*boxed_ptr)))
    }
}

impl From<*mut *mut IUnknown> for Variant {
    fn from(ptr: *mut *mut IUnknown) -> Variant {
        Variant::PUnknown(Box::new(unsafe { *ptr }))
    }
}

impl From<*mut *mut IDispatch> for Variant {
    fn from(ptr: *mut *mut IDispatch) -> Variant {
        Variant::PDispatch(Box::new(unsafe { *ptr }))
    }
}
