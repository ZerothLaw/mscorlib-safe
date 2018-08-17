// result.rs - MIT License
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

use std;
//use failure::Error;
use winapi::shared::winerror::HRESULT;

#[derive(Debug, Fail)]
pub enum CommonHresultValues {
    #[fail(display = "Operation successful")]
    SOk = 0x00000000,
    #[fail(display = "Operation aborted")] 
    EAbort = 0x80004004, 
    #[fail(display = "General access denied error")]
    EAccessDenied = 0x80070005, 
    #[fail(display = "Unspecified failure")]
    EFail = 0x80004005, 
    #[fail(display = "Invalid handle")]
    EHandle = 0x80070006, 
    #[fail(display = "One or more arguments are invalid")]
    EInvalidArg = 0x80070057,
    #[fail(display = "No such interface supported")] 
    ENoInterface = 0x80004002, 
    #[fail(display = "Not implemented")]
    ENotImpl = 0x80004001, 
    #[fail(display = "Failed to allocate necessary memory")]
    EOutOfMemory = 0x8007000E,
    #[fail(display = "Invalid pointer")]
    EPointer = 0x80004003,
    #[fail(display = "Unexpected failure")]
    EUnexpected = 0x8000FFFF, 
    #[fail(display = "Unknown HRESULT value")]
    EUnknownValue = 0xFFFFFFFF,
}


pub struct Hresult(HRESULT);

impl From<Hresult> for CommonHresultValues {
    fn from(hr: Hresult) -> CommonHresultValues{
        use CommonHresultValues::*;
        match hr.0 as u32{
            0x0u32 => SOk, 
            0x80004004u32 => EAbort, 
            0x80070005u32 => EAccessDenied, 
            0x80004005u32 => EFail, 
            0x80070006u32 => EHandle, 
            0x80070057u32 => EInvalidArg, 
            0x80004002u32 => ENoInterface, 
            0x80004001u32 => ENotImpl, 
            0x8007000Eu32 => EOutOfMemory, 
            0x80004003u32 => EPointer, 
            0x8000FFFFu32 => EUnexpected, 
            _ => EUnknownValue,
        }
    }
}

#[derive(Debug, Fail)]
pub enum SourceLocation {
    #[fail(display = "ICollection(line: {})", _0)]
    ICollection(u32),
    #[fail(display = "IComparable(line: {})", _0)]
    IComparable(u32),
    #[fail(display = "IComparer(line: {})", _0)]
    IComparer(u32), 
    #[fail(display = "IDictionary(line: {})", _0)]
    IDictionary(u32),
    #[fail(display = "IDictionaryEnumerator(line: {})", _0)]
    IDictionaryEnumerator(u32),
    #[fail(display = "IEnumerable(line: {})", _0)]
    IEnumerable(u32),
    #[fail(display = "IEnumerator(line: {})", _0)]
    IEnumerator(u32),
    #[fail(display = "IEqualityComparer(line: {})", _0)]
    IEqualityComparer(u32),
    #[fail(display = "IHashCodeProvider(line: {})", _0)]
    IHashCodeProvider(u32),
    #[fail(display = "IList(line: {})", _0)]
    IList(u32),
    #[fail(display = "Assembly(line: {})", _0)]
    _Assembly(u32),
    #[fail(display = "Type(line: {})", _0)]
    _Type(u32),
    #[fail(display = "MemberInfo(line: {})", _0)]
    _MemberInfo(u32),
    #[fail(display = "_MethodBase(line: {})", _0)]
    _MethodBase(u32),
    #[fail(display = "_MethodInfo(line: {})", _0)]
    _MethodInfo(u32),
    #[fail(display = "_ConstructorInfo(line: {})", _0)]
    _ConstructorInfo(u32),
    #[fail(display = "_FieldInfo(line: {})", _0)]
    _FieldInfo(u32),
    #[fail(display = "_PropertyInfo(line: {})", _0)]
    _PropertyInfo(u32),
    #[fail(display = "_EventInfo(line: {})", _0)]
    _EventInfo(u32),
}

#[derive(Debug, Fail)]
pub enum ClrError {
    #[fail(display = "unsafe call in {:?} resulted in a non-zero HRESULT: 0x{:x}", source, hr)]
    InnerCall{
        hr: HRESULT, 
        source: SourceLocation
    }, 
    #[fail(display = "Conversion failed at: {:?}", source)]
    Conversion{
        source: SourceLocation
    }, 
}

pub type Result<T> = std::result::Result<T, ClrError>;