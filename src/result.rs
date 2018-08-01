use std;
//use failure::Error;
use winapi::shared::winerror::HRESULT;

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