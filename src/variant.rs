
use std::mem;
use std::ops::Deref;
use std::marker::PhantomData;
use std::ptr;

use winapi::shared::wtypes::{VT_DISPATCH, VT_UNKNOWN, VARTYPE};
use winapi::shared::wtypes::{CY, DATE, VT_I8};
use winapi::shared::wtypesbase::SCODE;

use winapi::um::oaidl::{IDispatch, VARIANT, VARIANT_n3, __tagVARIANT, VARIANT_n1};
use winapi::um::unknwnbase::IUnknown;

use primitives::{Primitive, Primitives};
use wrappers::PtrContainer;

pub type LPVARIANT = *mut VARIANT;

pub enum Variant<VD, VU, TDispatch, TUnknown, TPrimitive>
    where VD: PtrContainer<TDispatch> , 
          TDispatch: Deref<Target=IDispatch>, 
          VU: PtrContainer<TUnknown> , 
          TUnknown: Deref<Target=IUnknown> 
{
    VariantDispatch(VD),
    VariantUnknown(VU), 
    VariantPrimitive(Box<Primitive<Target=TPrimitive>>),
    Phantom(PhantomData<TDispatch>, PhantomData<TUnknown>)
}

pub use Variant::*;

impl<VD, VU, TDispatch, TUnknown, TPrimitive> From<LPVARIANT> for Variant<VD, VU, TDispatch, TUnknown, TPrimitive> 
    where VD: PtrContainer<TDispatch> , 
          TDispatch: Deref<Target=IDispatch>, 
          VU: PtrContainer<TUnknown> , 
          TUnknown: Deref<Target=IUnknown>,
{
    fn from(vt: LPVARIANT) -> Variant<VD, VU, TDispatch, TUnknown, TPrimitive> {
        let mut n1 = unsafe{
            (*vt).n1
        };
        let n2_mut = unsafe {
            n1.n2_mut()
        };
        let vt: VARTYPE = n2_mut.vt;
        let mut n3 = n2_mut.n3;
        match vt as u32 {
            VT_DISPATCH => unsafe {
                let pn3 = n3.pdispVal_mut();
                let pn3 = mem::transmute::<*mut IDispatch, *mut TDispatch>(*pn3);
                let vd = VariantDispatch(VD::from(pn3));
                vd
            }, 
            VT_UNKNOWN => unsafe {
                let pn3 = n3.punkVal_mut();
                let pn3 = mem::transmute::<*mut IUnknown, *mut TUnknown>(*pn3);
                VariantUnknown(VU::from(pn3))
            }, 

            _=> Phantom(PhantomData, PhantomData)
        }

        //Phantom(PhantomData, PhantomData)
    }
}   

impl<VD, VU, TDispatch, TUnknown, TPrimitive> From<Variant<VD, VU, TDispatch, TUnknown, TPrimitive>> for VARIANT 
    where VD: PtrContainer<TDispatch>, 
          TDispatch: Deref<Target=IDispatch>, 
          VU: PtrContainer<TUnknown>, 
          TUnknown: Deref<Target=IUnknown> 
{
    fn from(variant: Variant<VD, VU, TDispatch, TUnknown, TPrimitive>) -> VARIANT {
        match variant {
            VariantDispatch(vd) => {
                let mut s = vd.ptr_mut();
                let n3 = unsafe {
                    let mut n3: VARIANT_n3 =  mem::zeroed();
                    {
                        let mut n_ptr = n3.pdispVal_mut();
                        *n_ptr = mem::transmute::<*mut TDispatch, *mut IDispatch>(s);
                    };
                    n3
                };
                let tv = __tagVARIANT { vt: VT_DISPATCH as u16, 
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
            }, 
            VariantUnknown(vu) => {
                let mut s = vu.ptr_mut();
                let n3 = unsafe {
                    let mut n3: VARIANT_n3 =  mem::zeroed();
                    {
                        let mut n_ptr = n3.punkVal_mut();
                        *n_ptr = mem::transmute::<*mut TUnknown, *mut IUnknown>(s);
                    };
                    n3
                };
                let tv = __tagVARIANT { vt: VT_UNKNOWN as u16, 
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
            }, 
            VariantPrimitive(vp) => {
                let vartype = vp.prim_type();
                let value = (*vp).get();
                let n3 = unsafe {
                    let mut n3: VARIANT_n3 =  mem::zeroed();
                    {
                        match vartype {
                            Primitives::I16 => {
                                let mut n_ptr = n3.iVal_mut();
                                *n_ptr = *mem::transmute::<&TPrimitive, &i16>(value);
                            }, 
                            Primitives::I32 => {
                                let mut n_ptr = n3.lVal_mut();
                                *n_ptr = *mem::transmute::<&TPrimitive, &i32>(value);
                            },
                            Primitives::F32 => {
                                let mut n_ptr = n3.fltVal_mut();
                                *n_ptr = *mem::transmute::<&TPrimitive, &f32>(value);
                            },
                            Primitives::F64 => {
                                let mut n_ptr = n3.dblVal_mut();
                                *n_ptr = *mem::transmute::<&TPrimitive, &f64>(value);
                            },
                            Primitives::Currency => {
                                let mut n_ptr = n3.cyVal_mut();
                                *n_ptr = *mem::transmute::<&TPrimitive, &CY>(value);
                            },
                            Primitives::Date => {
                                let mut n_ptr = n3.date_mut();
                                *n_ptr = *mem::transmute::<&TPrimitive, &DATE>(value);
                            },
                            Primitives::ErrorCode => {
                                let mut n_ptr = n3.scode_mut();
                                *n_ptr = *mem::transmute::<&TPrimitive, &SCODE>(value);
                            },
                            Primitives::Bool => {
                                let mut n_ptr = n3.boolVal_mut();
                                let b = *mem::transmute::<&TPrimitive, &bool>(value);
                                *n_ptr = if b {-1} else {0};
                            },
                            Primitives::Decimal => {
                                panic!("Can't safely consume DECIMAL into VARIANT as it takes a pointer, and nothing is moved.");
                            },
                            Primitives::Char => {
                                // let mut n_ptr = n3.boolVal_mut();
                                // let b = *mem::transmute::<&TPrimitive, &bool>(value);
                                // *n_ptr = if b {-1} else {0};
                            },
                            Primitives::UChar => {
                                // let mut n_ptr = n3.boolVal_mut();
                                // let b = *mem::transmute::<&TPrimitive, &bool>(value);
                                // *n_ptr = if b {-1} else {0};
                            },
                            Primitives::U16 => {
                                let mut n_ptr = n3.uiVal_mut();
                                *n_ptr = *mem::transmute::<&TPrimitive, &u16>(value);
                            },
                            Primitives::U32 => {
                                let mut n_ptr = n3.ulVal_mut();
                                *n_ptr =  *mem::transmute::<&TPrimitive, &u32>(value);
                            },
                            Primitives::I64 => {
                                let mut n_ptr = n3.llVal_mut();
                                *n_ptr =*mem::transmute::<&TPrimitive, &i64>(value);
                            },
                            Primitives::U64 => {
                                let mut n_ptr = n3.ullVal_mut();
                                *n_ptr = *mem::transmute::<&TPrimitive, &u64>(value);
                            },
                            Primitives::Void => {
                                // let mut n_ptr = n3.ullVal_mut();
                                // let b = *mem::transmute::<&TPrimitive, &u64>(value);
                                // *n_ptr = if b {-1} else {0};
                            },
                            Primitives::Ptr => {
                                // let mut n_ptr = n3.ullVal_mut();
                                // let b = *mem::transmute::<&TPrimitive, &u64>(value);
                                // *n_ptr = if b {-1} else {0};
                            },
                            Primitives::CNulTermString => {
                                // let mut n_ptr = n3.ullVal_mut();
                                // let b = *mem::transmute::<&TPrimitive, &u64>(value);
                                // *n_ptr = if b {-1} else {0};
                            },
                            Primitives::CWideNulTermString => {
                                // let mut n_ptr = n3.ullVal_mut();
                                // let b = *mem::transmute::<&TPrimitive, &u64>(value);
                                // *n_ptr = if b {-1} else {0};
                            },
                            Primitives::IntPtr => {
                                // let mut n_ptr = n3.ullVal_mut();
                                // let b = *mem::transmute::<&TPrimitive, &u64>(value);
                                // *n_ptr = if b {-1} else {0};
                            },
                            Primitives::UIntPtr => {
                                // let mut n_ptr = n3.ullVal_mut();
                                // let b = *mem::transmute::<&TPrimitive, &u64>(value);
                                // *n_ptr = if b {-1} else {0};
                            },
                            Primitives::ClassId => {
                                // let mut n_ptr = n3.ullVal_mut();
                                // let b = *mem::transmute::<&TPrimitive, &u64>(value);
                                // *n_ptr = if b {-1} else {0};
                            },
                            Primitives::ByRef => {
                                // let mut n_ptr = n3.ullVal_mut();
                                // let b = *mem::transmute::<&TPrimitive, &u64>(value);
                                // *n_ptr = if b {-1} else {0};
                            },
                        }
                    };
                    n3
                };
                let tv = __tagVARIANT { vt: VARTYPE::from(vartype) as u16, 
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
            Phantom(_, _) => {
                unreachable!()
            }
        }
    }
}

pub struct PhantomDispatch {}
pub struct PhantomUnknown {}

pub struct WrappedDispatch {
    ptr: *mut PhantomDispatch
}

pub struct WrappedUnknown {
    ptr: *mut PhantomUnknown
}

impl Deref for PhantomDispatch {
    type Target = IDispatch;
    fn deref(&self) -> &Self::Target {
        unsafe {&*(ptr::null_mut())}
    }
}

impl Deref for PhantomUnknown {
    type Target = IUnknown;
    fn deref(&self) -> &Self::Target {
        unsafe{&*ptr::null_mut()}
    }
}

impl PtrContainer<PhantomDispatch> for WrappedDispatch {
    fn ptr(&self) -> *const PhantomDispatch {
        self.ptr
    }
    fn ptr_mut(&self) -> *mut PhantomDispatch {
        self.ptr
    }

    fn from(ppd: *mut PhantomDispatch) -> WrappedDispatch {
        WrappedDispatch { ptr: ppd}
    }
}

impl PtrContainer<PhantomUnknown> for WrappedUnknown {
    fn ptr(&self) -> *const PhantomUnknown {
        self.ptr
    }
    fn ptr_mut(&self) -> *mut PhantomUnknown {
        self.ptr
    }
    
    fn from(ppu: *mut PhantomUnknown) -> WrappedUnknown {
        WrappedUnknown{ptr: ppu}
    }
}

pub type DispatchVariant<VDispatch, TDispatch> = Variant<VDispatch, WrappedUnknown, TDispatch, PhantomUnknown, i16>;
pub type UnknownVariant<VUnknown, TUnknown> = Variant<WrappedDispatch, VUnknown, PhantomDispatch, TUnknown, i16>;
pub type PrimitiveVariant<TPrimitive> = Variant<WrappedDispatch, WrappedUnknown, PhantomDispatch, PhantomUnknown, TPrimitive>;

//LPVARIANT::from(Variant::VariantDispatch::<R, WrappedUnknown, IComparable, PhantomUnknown>(rhs));
//R, IComparable, rhs
macro_rules! DISPATCH {
    ($var:ident : $t:ty : $wt:ty ) => {
        VARIANT::from(Variant::VariantDispatch::<$t, WrappedUnknown, $wt, PhantomUnknown, i16>($var))
    };
}

macro_rules! UNKNOWN {
    ($var:ident : $t:ty : $wt:ty ) => {
        VARIANT::from(Variant::VariantUnknown::< WrappedDispatch, $t, PhantomDispatch, $wt, i16>($var))
    };
}