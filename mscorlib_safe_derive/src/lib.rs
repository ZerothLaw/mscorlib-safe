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
                let gen_t_ptr = tptr; //raw pointer, we need elem as well
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
                    }
                };

                return expanded.into()
            }
        }
    }
    let e = quote!{};

    return e.into();
}