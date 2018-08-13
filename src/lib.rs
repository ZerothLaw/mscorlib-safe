#![feature(never_type)]

#[macro_use] extern crate failure;

extern crate winapi;

extern crate mscorlib_sys;

#[macro_use]pub mod macros;

extern crate rust_decimal;

mod bstring;
mod collections;
mod params;
mod result;
mod struct_wrappers;
mod wrappers;

pub mod new_variant;
pub mod new_safearray;

pub use collections::*;
pub use bstring::*;
pub use result::*;
pub use wrappers::*;

use mscorlib_sys::system::reflection::{_Assembly, _AssemblyName, _Binder, _ConstructorInfo, _FieldInfo, _EventInfo, _MemberInfo, 
_MethodBase, _Module, _ParameterInfo, _PropertyInfo, _Type};
use mscorlib_sys::system::collections::{ICollection, IComparer, IDictionary, IDictionaryEnumerator, 
IEnumerable, IEnumerator, IEqualityComparer, IHashCodeProvider, IList};
use mscorlib_sys::system::{IComparable};

macro_rules! BLANKET_IMPLS {
    ($({$tr:ty, $ptr_ty:ty},)*) => {
        $(
            impl<T:PtrContainer<$ptr_ty>> $tr for T{}
        )*
    };
}

BLANKET_IMPLS!{
    {Collection, ICollection},
    {Comparable, IComparable},
    {Comparer, IComparer},
    {Dictionary, IDictionary},
    {DictionaryEnumerator, IDictionaryEnumerator},
    {Enumerable, IEnumerable}, 
    {Enumerator, IEnumerator},
    {EqualityComparer, IEqualityComparer}, 
    {HashCodeProvider, IHashCodeProvider}, 
    {List, IList}, 
    {Assembly, _Assembly}, 
    {Type, _Type}, 
    {MemberInfo, _MemberInfo}, 
    {MethodBase, _MethodBase}, 
    {ConstructorInfo, _ConstructorInfo}, 
    {FieldInfo, _FieldInfo}, 
    {PropertyInfo, _PropertyInfo}, 
    {EventInfo, _EventInfo}, 
    {ParameterInfo, _ParameterInfo}, 
    {Module, _Module}, 
    {AssemblyName, _AssemblyName},
    {Binder, _Binder},
}