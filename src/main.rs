use anyhow::Error;
use std::path::Path;
use pdf_extract::extract_text;
use rust_xlsxwriter::{cell_range, Format, Formula, Workbook};
use core::f64;
use std::fs;
use regex::Regex;

fn main() -> Result<(), Error> {

    let args: Vec<String> = std::env::args().collect();
    let path = Path::new(&args[0]);
    std::env::set_current_dir(path.parent().unwrap())?;
    
    let pdf_dir = "pdf/";

    let pdf_paths: Vec<_> = fs::read_dir(pdf_dir)?
        .filter_map(Result::ok)
        .filter(|entry| entry.path().extension().map_or(false, |ext| ext == "pdf"))
        .map(|entry| entry.path())
        .collect();
    /*
     * The code below is for extracting text and writing it to an Excel file
     */

    //* safe the extracted text in a vector
    let mut extracted_texts: Vec<(String, Vec<String>)> = Vec::new();

    for pdf_path in pdf_paths {
        match extract_text(&pdf_path) {
            Ok(text) => {
                //* println!("Extrahierter Text von {}: {}", pdf_path.display(), text);
                //* Teile den Text in verschiedene Zeilen auf, filtere leere Zeilen und speichere sie im Vektor
                let split_texts: Vec<String> = text
                    .split('\n')
                    .map(|s| s.to_string())
                    .filter(|s| !s.trim().is_empty())
                    .collect();
                extracted_texts.push((pdf_path.to_str().unwrap().to_string(), split_texts));
            }
            Err(err) => eprintln!(
                "Fehler beim Extrahieren von {}: {}",
                pdf_path.display(),
                err
            ),
        }
    }

    //* create a new Excel doc
    let mut workbook = Workbook::new();

    //* Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    //* For clarity, define some variables to use in the formula and chart ranges.
    //* Row and column numbers are all zero-indexed.
    let first_row = 2; // Skip the header row.
    let last_row = first_row + (extracted_texts.len() as u32) - 1;
    let netto_col = 3;
    let brutto_col = 5;

    let netto_range = cell_range(first_row, netto_col, last_row, netto_col);
    let netto_formula = format!("=SUM({})", netto_range);
    let brutto_range = cell_range(first_row, brutto_col, last_row, brutto_col);
    let brutto_formula = format!("=SUM({})", brutto_range);

    //* Text Bold
    let bold = Format::new().set_bold();

    //* Write Headers
    worksheet.write_with_format(0, 0, "PDF File Name", &bold)?;
    worksheet.write_with_format(0, 1, "Belegnummer", &bold)?;
    worksheet.write_with_format(0, 2, "Datum", &bold)?;
    worksheet.write_with_format(0, 3, "Zwischensumme (Netto)", &bold)?;
    worksheet.write_with_format(1, 3, Formula::new(netto_formula), &bold)?;
    worksheet.write_with_format(0, 4, "USt", &bold)?;
    worksheet.write_with_format(0, 5, "Zwischensumme (Brutto)", &bold)?;
    worksheet.write_with_format(1, 5, Formula::new(brutto_formula), &bold)?;

    //* Set column width
    worksheet.set_column_width(0, 15)?;
    worksheet.set_column_width(1, 30)?;
    worksheet.set_column_width(2, 25)?;
    worksheet.set_column_width(3, 35)?;
    worksheet.set_column_width(4, 15)?;
    worksheet.set_column_width(5, 35)?;

//* Extract text from each PDF and store it in the vector
let mut row = 2;
let re = Regex::new(r"(\d+\.\d+)").unwrap();
for (pdf_file, texts) in extracted_texts {
    worksheet.write(row, 0, &pdf_file)?;

    let mut receipt_number = String::new();
    let mut date = String::new();
    let mut netto_zwi = 0.0;
    let mut brutto_zwi = 0.0;
    let mut ust = String::new();


    for text in texts {
        if text.contains("Belegnummer") || text.contains("Receipt number") {
            receipt_number = text.clone();
        }
        if text.contains("Datum") || text.contains("Date") {
            date = text.clone();
        }
        if text.contains("Zwischensumme") || text.contains("Sub total") {
            if let Some(caps) = re.captures(&text) {
                netto_zwi = caps.get(0).map_or("NULL", |m| m.as_str()).parse::<f64>().unwrap_or(0.0);
            }
        }
        if text.contains("USt. 0.00%") || text.contains("VAT 0.00%") {
            ust = text.clone();
        }
        if text.contains("Gesamtbetrag") || text.contains("Total") {
            if let Some(caps) = re.captures(&text) {
                brutto_zwi = caps.get(0).map_or("NULL", |m| m.as_str()).parse::<f64>().unwrap_or(0.0);
            }
        }
    }

    worksheet.write(row, 1, &receipt_number)?;
    worksheet.write(row, 2, &date)?;
    worksheet.write(row, 3, netto_zwi)?;
    worksheet.write(row, 4, &ust)?;
    worksheet.write(row, 5, brutto_zwi)?;

    row += 1;
}

    //* Konvertiere die Werte in den Spalten 15 und 17 in Ganzzahlen

    //* Save the file to disk.
    workbook.save("Test.xlsx")?;
    println!("Excel-Datei wurde erfolgreich erstellt");

    Ok(())
}
