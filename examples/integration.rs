// integration.rs - MIT License
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

extern crate winapi;

extern crate mscorlib_sys;
extern crate mscorlib_safe;

#[macro_use]
extern crate mscorlib_safe_derive;

use winapi::um::oaidl::{LPSAFEARRAY};

use mscorlib_sys::system::reflection::_Type;
use mscorlib_safe::PtrContainer;
use mscorlib_safe::new_variant::Variant;
use mscorlib_safe::new_safearray::RSafeArray;

#[derive(Debug,PtrContainer)]
struct Fit {
    ptr: *mut _Type
}

fn main() {
    let mut vc = Vec::new();
    vc.push(100u32);

    let rsa: RSafeArray<u32> = RSafeArray::from(vc);
    let _psa = LPSAFEARRAY::from(rsa);
}