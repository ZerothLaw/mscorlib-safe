use winapi::ctypes::{c_void};
use winapi::shared::guiddef::CLSID;

use winapi::shared::wtypes::{CY,
                             VARTYPE,    VT_BOOL,  VT_BYREF,    VT_CLSID, 
                             VT_CY,      VT_DATE,  VT_DECIMAL,  VT_ERROR,  
                             VT_INT_PTR, VT_I1,    VT_I2,       VT_I4, 
                             VT_I8,      VT_LPSTR, VT_LPWSTR,   VT_PTR,      
                             VT_R4,      VT_R8,    VT_UINT_PTR, VT_UI1, 
                             VT_UI2,     VT_UI4,   VT_UI8,      VT_VOID, };

pub enum Primitives {
    I16 , 
    I32 , 
    F32 , 
    F64 , 
    Currency , 
    Date , 
    ErrorCode , 
    Bool , 
    Decimal , 
    Char , 
    UChar , 
    U16 , 
    U32 , 
    I64 , 
    U64 , 
    Void , 
    Ptr ,
    CNulTermString , 
    CWideNulTermString , 
    IntPtr , 
    UIntPtr , 
    ClassId ,
    ByRef,
}

impl From<Primitives> for VARTYPE {
    fn from(p: Primitives) -> VARTYPE {
        use Primitives::*;
        let t = match p {
            I16 => VT_I2, 
            I32 => VT_I4, 
            F32 => VT_R4, 
            F64 => VT_R8, 
            Currency => VT_CY, 
            Date => VT_DATE, 
            ErrorCode => VT_ERROR, 
            Bool => VT_BOOL, 
            Decimal => VT_DECIMAL, 
            Char => VT_I1, 
            UChar => VT_UI1, 
            U16 => VT_UI2, 
            U32 => VT_UI4, 
            I64 => VT_I8, 
            U64 => VT_UI8, 
            Void => VT_VOID, 
            Ptr => VT_PTR,
            CNulTermString => VT_LPSTR, 
            CWideNulTermString => VT_LPWSTR, 
            IntPtr => VT_INT_PTR, 
            UIntPtr => VT_UINT_PTR, 
            ClassId => VT_CLSID,
            ByRef => VT_BYREF
        };
        t as u16
    }
}

pub trait Primitive {
    type Target;
    fn get(&self) -> &Self::Target;
    fn prim_type(&self) -> Primitives;
}

macro_rules! WRAP_IMPLS {
    ($({$p_type:ty, $e_type:ident})*) => {
        $(
            impl Primitive for $p_type {
                type Target = $p_type;
                fn get(&self) -> &Self::Target {
                    self
                }
                fn prim_type(&self) -> Primitives {
                    Primitives::$e_type
                }
            }
        )*
    };
}

WRAP_IMPLS!{
    {i16, I16}
    {i32, I32}
    {f32, F32}
    {f64, F64}
    {CY, Currency}
    {u16, U16}
    {u32, U32}
    {i64, I64}
    {u64, U64}
    {c_void, Void}
    {CLSID, ClassId}
}