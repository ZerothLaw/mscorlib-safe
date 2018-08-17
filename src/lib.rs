#![feature(never_type)]
// lib.rs - MIT License
//  Copyright (c) 2018 Tyler Laing (ZerothLaw)
// 
//  Permission is hereby granted, free of charge, to any person obtaining a copy
//  of this software and associated documentation files (the "Software"), to deal
//  in the Software without restriction, including without limitation the rights
//  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
//  copies of the Software, and to permit persons to whom the Software is
//  furnished to do so, subject to the following conditions:
// 
//  The above copyright notice and this permission notice shall be included in all
//  copies or substantial portions of the Software.
// 
//  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
//  SOFTWARE.

#[macro_use] extern crate failure;

extern crate winapi;

extern crate mscorlib_sys;

#[macro_use]pub mod macros;

extern crate rust_decimal;

mod bstring;
mod collections;
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