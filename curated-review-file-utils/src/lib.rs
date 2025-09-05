//! Library for parsing curated review Excel files
//! Provides utilities to extract header information
//! and generate reports and YAML lookups.

use calamine::{open_workbook, Reader, Xlsx};
use chrono::Local;
use serde::Serialize;
use std::fs::{self, File};
use std::io::{self, Write};
use std::path::{Path, PathBuf};
use whoami;

#[derive(Debug, Serialize)]
pub struct HeaderLookups {
    pub column_name_lookup: Vec<(String, usize)>,
    pub index_position_lookup: Vec<(usize, String)>,
}

/// Parse headers from the first worksheet of the Excel file.
pub fn parse_headers<P: AsRef<Path>>(infile: P) -> io::Result<Vec<String>> {
    let mut workbook: Xlsx<_> = open_workbook(infile)?;
    let range = workbook.worksheet_range_at(0).unwrap().unwrap();

    // Find the "real" header row
    let mut rows = range.rows();
    let first_row = rows.next().unwrap_or(&[]);
    let header_row = if first_row[0].to_string().contains("Labnumber") {
        first_row
    } else {
        rows.next().unwrap_or(&[])
    };

    Ok(header_row.iter().map(|c| c.to_string()).collect())
}

/// Write the report file.
pub fn write_report<P: AsRef<Path>>(
    report_path: P,
    infile: &Path,
    headers: &[String],
    log_path: &Path,
    exe_path: &Path,
) -> io::Result<()> {
    let mut file = File::create(&report_path)?;
    let now = Local::now();

    writeln!(file, "## method-created: {}", exe_path.display())?;
    writeln!(file, "## date-created: {}", now.format("%Y-%m-%d %H:%M:%S"))?;
    writeln!(file, "## created-by: {}", whoami::username())?;
    writeln!(file, "## infile: {}", infile.display())?;
    writeln!(file, "## report-file: {}", report_path.as_ref().display())?;
    writeln!(file, "## logfile: {}", log_path.display())?;
    writeln!(file)?;

    // Section 1
    writeln!(file, "# Section 1: Column Name -> Index")?;
    for (i, col) in headers.iter().enumerate() {
        writeln!(file, "{}\t{}", col, i)?;
    }

    writeln!(file)?;
    writeln!(file, "# Section 2: Index -> Column Name")?;
    for (i, col) in headers.iter().enumerate() {
        writeln!(file, "{}\t{}", i, col)?;
    }

    Ok(())
}

/// Write YAML file with header lookups.
pub fn write_yaml<P: AsRef<Path>>(yaml_path: P, headers: &[String]) -> io::Result<()> {
    let lookups = HeaderLookups {
        column_name_lookup: headers.iter().enumerate().map(|(i, c)| (c.clone(), i)).collect(),
        index_position_lookup: headers.iter().enumerate().map(|(i, c)| (i, c.clone())).collect(),
    };
    let yaml_str = serde_yaml::to_string(&lookups).unwrap();
    fs::write(yaml_path, yaml_str)
}
