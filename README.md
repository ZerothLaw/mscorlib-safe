# mscorlib-safe
Safe wrappers around raw FFI bindings in mscorlib-sys.

Provides traits and structs for interacting with FFI bindings in Rusty ways. 

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. See deployment for notes on how to deploy the project on a live system.

### Prerequisites

What things you need to install the software and how to install them

```
Windows 10
Visual Studio Community 2017
.Net >= 4.6.1
Rust >= 1.28
```

These are needed for mscorlib-sys, not mscorlib-safe.

### Installing

Clone the repository

```
git clone https://github.com/ZerothLaw/mscorlib-safe.git
```

Move into the directory

```
cd mscorlib-safe/
```

Build the library

```
cargo build
```

See examples\integration.rs for an example of how to work with the wrappers.

Or add this line to your .toml file:

```
[dependencies]
mscorlib-safe = "0.1.3"
```

## Running the tests

```
cargo test
```

## Deployment

This is intended to be compiled into a final executable, not deployed on its own. 

## Built With

* [Rust](https://www.rust-lang.org/) - the language, compiler, and package management
* [winapi-rs](https://github.com/retep998/winapi-rs) - invaluable model for the bindings, as well as original developer of many of the macros used.

## Contributing

Please read [CONTRIBUTING.md]() for details on our code of conduct, and the process for submitting pull requests to us.

## Versioning

We use [SemVer](http://semver.org/) for versioning. For the versions available, see the [tags on this repository](https://github.com/ZerothLaw/mscorlib-safe/tags). 

## Authors

* **Tyler Laing** - *Initial work* - [ZerothLaw](https://github.com/ZerothLaw)

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

* Rust developers and designers for an excellent language and environment
* retep998 for winapi
