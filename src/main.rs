use clap::Parser;
use std::io::prelude::*;
use std::path::{Path, PathBuf};
use std::{env, fs, io};
use zip::result::ZipResult;
use zip::{ZipArchive, ZipWriter};

#[derive(Parser)]
#[clap(name = "pwsh-b64-rs")]
#[clap(author = "Petruknisme <me@petruknisme.com>")]
#[clap(version = "1.0")]
#[clap(about = "Powershell implementation of ToBase64String UTF-16LE written in Rust", long_about = None)]

struct Cli {
    /// Template URL
    #[clap(short, long)]
    url: String,

    /// File to be injected
    #[clap(short, long)]
    file: String,
}

fn unzip(zipfile: &str, dest_dir: &str) -> ZipResult<()> {
    let file = fs::File::open(zipfile)?;
    let mut archive = ZipArchive::new(file)?;

    for i in 0..archive.len() {
        let mut file = archive.by_index(i)?;
        let filename = file.name();
        if filename.starts_with("__MACOSX") {
            continue;
        }
        let outpath = Path::new(dest_dir).join(filename);

        if (file.name()).ends_with('/') {
            fs::create_dir_all(&outpath)?;
        } else {
            if let Some(p) = outpath.parent() {
                if !p.exists() {
                    fs::create_dir_all(p)?;
                }
            }
            let mut outfile = fs::File::create(&outpath)?;
            io::copy(&mut file, &mut outfile)?;
        }
    }

    Ok(())
}

fn zip_dir(dir: &Path, zip: &mut ZipWriter<std::fs::File>, prefix: &Path) -> ZipResult<()> {
    let files = fs::read_dir(dir)?;

    for file in files {
        let file = file?;
        let path = file.path();
        let name = path.strip_prefix(prefix).unwrap();

        if path.is_file() {
            zip.start_file(name.to_str().unwrap(), zip::write::FileOptions::default())?;
            let mut f = fs::File::open(path)?;
            let mut buffer = Vec::new();
            f.read_to_end(&mut buffer)?;
            zip.write_all(&buffer)?;
        } else if path.is_dir() {
            zip_dir(&path, zip, prefix)?;
        }
    }

    Ok(())
}

fn main() {
    //let current_dir = env::current_dir().unwrap();
    //let current_dir_path = current_dir.as_path();

     let zipfile = "/tmp/test2/twrp.zip";
    let dest_dir = "/tmp/test2/";

    unzip(zipfile, dest_dir).unwrap(); 

    /* let dir = Path::new("/tmp/test2/twrp");
    let zipfile = "/tmp/test2/twrp2.zip";

    let prefix = dir.parent().unwrap();

    let file = fs::File::create(zipfile).unwrap();
    let mut zip = ZipWriter::new(file);

    zip_dir(dir, &mut zip, prefix).unwrap();
    zip.finish().unwrap(); */
}
