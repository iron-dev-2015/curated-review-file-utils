# Curated Review File Utils

This workspace contains:

- **curated-review-file-utils**  
  A Rust library for parsing curated review Excel `.xlsx` files.  
  Provides functions for extracting headers, generating report files, and YAML lookup tables.

- **get-curated-review-xlsx-headers**  
  A CLI tool that wraps the library and outputs headers to reports, YAML, and logs.

## Usage

### Library
Add to `Cargo.toml`:
```toml
curated-review-file-utils = "0.1"
