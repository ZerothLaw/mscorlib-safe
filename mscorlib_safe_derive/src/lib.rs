#![recursion_limit="128"]
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

extern crate proc_macro;
extern crate proc_macro2;
extern crate syn;
#[macro_use]
extern crate quote;

use proc_macro2::{Ident, Span};
use syn::{DeriveInput, Data, Type};

// impl PtrContainer<_Type> for ClrType {
//     fn ptr(&self) -> *const _Type {
//         self.ptr
//     }

//     fn ptr_mut(&self) -> *mut _Type {
//         self.ptr
//     }
// }

#[proc_macro_derive(PtrContainer)]
pub fn pointer_container_derive(input: proc_macro::TokenStream) -> proc_macro::TokenStream {
    let inp: DeriveInput = syn::parse(input).unwrap();
    let name = inp.ident;
    if let Data::Struct(dta) = inp.data {
        for (ix, field) in dta.fields.iter().enumerate() { 
            if let Type::Ptr( ref tptr) = &field.ty {
                let f_name = match field.ident {
                    Some(ref fnn) => fnn.to_string(), 
                    None => ix.to_string()
                };
                let f_name = Ident::new(&f_name, Span::call_site());
                //let f_name = f_name.unwrap();
                let elem = &tptr.elem;
                let mut expanded = quote! {
                    //impl ToVariantUnknown<_Type> for ClrType {} 
                    impl PtrContainer<#elem> for #name {
                        fn ptr(&self) -> *const #elem {
                            self.#f_name
                        }
                        fn ptr_mut(&self) -> *mut #elem {
                            self.#f_name
                        }
                        fn from(p: *mut #elem) -> #name {
                            #name {#f_name: p}
                        }
                        fn into_variant(&self) -> Variant {
                            use std::mem;
                            use winapi::um::unknwnbase::IUnknown;
                            let p = self.ptr_mut();
                            let p: *mut IUnknown = unsafe {mem::transmute::<*mut #elem, *mut IUnknown>(p)};
                            Variant::from(p)
                        }
                    }
                };

                return expanded.into()
            }
        }
    }
    let e = quote!{};

    return e.into();
}