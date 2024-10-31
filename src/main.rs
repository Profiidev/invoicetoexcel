use std::{f64::consts::E, fs};

use lopdf::{content, Document};
use pdf_extract::extract_text;
use anyhow::Error;

fn main() -> Result<(), Error> {
    
    let pdf_dir = "pdf/"; 

    let pdf_paths: Vec<_> = fs::read_dir(pdf_dir)?
        .filter_map(Result::ok)
        .filter(|entry| entry.path().extension().map_or(false, |ext| ext == "pdf"))
        .map(|entry| entry.path())
        .collect();

        for pdf_path in pdf_paths {
            match extract_text(&pdf_path) {
                Ok(text) => {
                    println!("Extrahierter Text von {}: {}", pdf_path.display(), text);
                    // You can use the `text` variable here for further processing
                },
                Err(err) => eprintln!("Fehler beim Extrahieren von {}: {}", pdf_path.display(), err),
            }
        }
        /*
         *  The code below is for extracting text from a single PDF file
         */
    //let pdf_path = "pdf/Rechnung.pdf"; // Hardcoded path to the PDF file

    //let text = extract_text(pdf_path)?;

    //println!("Extrahierter Text:\n{}", text);

    Ok(())

}
