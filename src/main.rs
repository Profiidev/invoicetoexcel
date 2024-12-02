use anyhow::Error;
use pdf_extract::extract_text;
use rust_xlsxwriter::{cell_range, Format, Formula, Workbook};
use std::{fs};

fn main() -> Result<(), Error> {
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
    let netto_col = 15;
    let brutto_col = 17;

    let netto_range = cell_range(first_row, netto_col, last_row, netto_col);
    let netto_formula = format!("=SUM({})", netto_range);
    let brutto_range = cell_range(first_row, brutto_col, last_row, brutto_col);
    let brutto_formula = format!("=SUM({})", brutto_range);

    //* Text Bold
    let bold = Format::new().set_bold();

    //* Write Headers
    worksheet.write_with_format(0, 0, "PDF File Name", &bold)?;
    worksheet.write_with_format(0, 1, "Belegnummer", &bold)?;
    worksheet.write_with_format(0, 5, "Datum", &bold)?;
    worksheet.write_with_format(0, 15, "Zwischensumme (Netto)", &bold)?;
    worksheet.write_with_format(1, 15, Formula::new(netto_formula), &bold)?;
    worksheet.write_with_format(0, 17, "Zwischensumme (Brutto)", &bold)?;
    worksheet.write_with_format(1, 17, Formula::new(brutto_formula), &bold)?;

    //* Set column width
    worksheet.set_column_width(0, 15)?;
    worksheet.set_column_width(1, 30)?;
    worksheet.set_column_width(5, 25)?;
    worksheet.set_column_width(15, 35)?;
    worksheet.set_column_width(17, 35)?;

//* Extract text from each PDF and store it in the vector
let mut row = 2;
for (pdf_file, texts) in extracted_texts {
    worksheet.write(row, 0, &pdf_file)?;
    for (col, text) in texts.iter().enumerate() {
        let col_index = (col + 1) as u32;
        if col_index == 15 || col_index == 17 {
            // Debugging-Ausgabe
            println!("Verarbeite Spalte {} in Zeile {} mit Wert: {}", col_index, row, text);
    
            // Entferne alle nicht-numerischen Zeichen außer Ziffern und Punkten
            let mut value_str: String = text.chars().filter(|c| c.is_digit(10) || *c == '.').collect();
            println!("Gefilterter Wert: {}", value_str); // Debugging-Ausgabe
    
            // Entferne vorangestellte Punkte
            while value_str.starts_with('.') {
                value_str.remove(0);
            }
            println!("Wert nach Entfernen der vorangestellten Punkte: {}", value_str); // Debugging-Ausgabe
    
            // Überprüfe, ob der String leer ist, und setze ihn ggf. auf "0"
            if value_str.is_empty() {
                value_str = "0".to_string();
            }
    
            // Konvertiere den bereinigten String in eine Zahl
            let value = value_str.parse::<f64>().unwrap_or(0.0);
            println!("Wert nach Parsen: {}", value); // Debugging-Ausgabe
    
            // Schreibe die Zahl in die Zelle
            worksheet.write_number(row, col_index.try_into().unwrap(), value)?;
        } else {
            worksheet.write(row, col_index.try_into().unwrap(), text)?;
        }
    }
    
    row += 1;
}

    //* Konvertiere die Werte in den Spalten 15 und 17 in Ganzzahlen

    //* Save the file to disk.
    workbook.save("Test.xlsx")?;
    println!("Excel-Datei wurde erfolgreich erstellt");

    Ok(())
}
