use std::slice;
use std::ffi::{OsString};
use std::ffi::OsStr;
use std::os::windows::ffi::{OsStrExt, OsStringExt};

use winapi::shared::minwindef::UINT;
use winapi::um::oleauto::{SysAllocStringLen, SysStringLen};

use wrappers::PtrContainer;

#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct BString {
	size: usize, 
	inner: Vec<u16>
}

#[derive(Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct BStr {
	size: usize,
	inner: [u16]
}

pub fn os_to_wide(s: &OsStr) -> Vec<u16> {
	s.encode_wide().collect()
}

pub fn os_from_wide(s: &[u16]) -> OsString {
	OsString::from_wide(s)
}

impl BString {
	pub fn new() -> BString {
		BString {inner: vec![], size: 0}
	}
	pub fn from_vec<T>(raw: T) -> BString 
		where T: Into<Vec<u16>>
	{
		let raw = raw.into();
		let l = raw.len();
		BString { inner: raw, size: l }
	}

	pub fn from_str<S>(s: &S) -> BString 
		where S: AsRef<OsStr> + ?Sized 
	{
		let s = os_to_wide(s.as_ref());
		let l = s.len();
		BString {
			inner: s, 
			size: l
		}
	}

	pub fn into_vec(self) -> Vec<u16> {
		self.inner
	}

	pub fn to_string(&self) -> String {
		String::from_utf16_lossy(&self.inner[..])
	}

	pub unsafe fn as_sys(&self) -> *mut u16 {
		let new_ws = self.inner.clone();
		let ws_ptr = new_ws.as_ptr();
		SysAllocStringLen(ws_ptr, self.size as UINT)
	}

	pub unsafe fn from_ptr(p: *const u16, len: usize) -> BString {
		if len == 0 {
			return BString::new();
		}
		assert!(!p.is_null());

		let slice = slice::from_raw_parts(p, len);
		BString::from_vec(slice)
	}

	pub fn from_ptr_safe(p: *mut u16) -> BString {
		assert!(!p.is_null());
		let us: u32 = unsafe {SysStringLen(p)};
		if us == 0 {
			return BString::new();
		}
		unsafe {BString::from_ptr(p, us as usize)}
	}
}

impl Into<Vec<u16>> for BString {
	fn into(self) -> Vec<u16> {
		self.into_vec()
	}
}

impl From<String> for BString {
	fn from(s: String) -> BString {
		BString::from_str(&s)
	}
}

impl From<&'static str> for BString {
	fn from(s: &str) -> BString {
		BString::from_str(s)
	}
}

impl From<BString> for String {
	fn from(bs: BString) -> String {
		bs.to_string()
	}
}

impl From<BString> for *mut u16 {
	fn from(bs: BString) -> *mut u16 {
		unsafe {bs.as_sys()}
	}
}

impl PtrContainer<u16> for BString {
	fn ptr(&self) -> *const u16 {
		unsafe {
			self.as_sys()
		}
	}
	fn ptr_mut(&self) -> *mut u16 {
		unsafe {
			self.as_sys()
		}
	}
	fn from(pmu: *mut u16) -> BString {
		BString::from_ptr_safe(pmu)
	}
}