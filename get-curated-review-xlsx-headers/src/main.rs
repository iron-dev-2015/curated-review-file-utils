use clap::Parser;
use curated_review_file_utils::{parse_headers, write_report, write_yaml};
use log::{error, info};
use std::env;
use std::fs;
use std::path::{Path, PathBuf};
use chrono::Local;

/// CLI tool wrapping curated-review-file-utils
#[derive(Parser)]
#[command(name = "get-curated-review-xlsx-headers", version, about)]
struct Cli {
    /// Input Excel file
    #[arg(long)]
    infile: PathBuf,

    /// Output directory
    #[arg(long)]
    outdir: Option<PathBuf>,

    /// Optional log file
    #[arg(long)]
    logfile: Option<PathBuf>,

    /// Optional report file
    #[arg(long)]
    report_file: Option<PathBuf>,
}

fn main() {
    env_logger::init();
    let args = Cli::parse();

    let user = whoami::username();
    let now = Local::now().format("%Y-%m-%d-%H%M%S").to_string();
    let exe_path = env::current_exe().unwrap();
    let exe_stem = exe_path.file_stem().unwrap().to_str().unwrap();

    let outdir = args.outdir.unwrap_or_else(|| {
        PathBuf::from(format!("/tmp/{}/curated-review-file-utils/{}", user, now))
    });
    fs::create_dir_all(&outdir).unwrap();

    let logfile = args.logfile.unwrap_or_else(|| outdir.join(format!("{}.log", exe_stem)));
    let report_file = args.report_file.unwrap_or_else(|| outdir.join(format!("{}_report.txt", exe_stem)));
    let yaml_file = outdir.join("column_headers.yaml");

    match parse_headers(&args.infile) {
        Ok(headers) => {
            if let Err(e) = write_report(&report_file, &args.infile, &headers, &logfile, &exe_path) {
                error!("Failed to write report: {}", e);
            }
            if let Err(e) = write_yaml(&yaml_file, &headers) {
                error!("Failed to write yaml: {}", e);
            }
            println!("{}", logfile.display());
            println!("{}", report_file.display());
            println!("{}", yaml_file.display());
        }
        Err(e) => {
            error!("Failed to parse headers: {}", e);
        }
    }
}
