use umya_spreadsheet::*;

pub struct CompanyDay {
    pub company: String,
    pub date: String,
    pub from: u32,
    pub to: u32,
    pub max: u32,
}

pub fn get_ranges(path_to_file: &std::path::PathBuf) -> Vec<CompanyDay> {
    let mut out: Vec<CompanyDay> = Vec::new();

    let book = reader::xlsx::read(path_to_file).unwrap();

    let mut row_idx = 2;

    let sheet = book.get_sheet_by_name(&"Sheet_1").unwrap();
    let mut cursor = CompanyDay {
        company: sheet.get_value((2, row_idx)),
        date: sheet.get_value((1, row_idx)),
        from: row_idx,
        to: row_idx,
        max: 0,
    };
    while row_idx < 10_000_000 {
        row_idx += 1;
        let date = sheet.get_value((1, row_idx));
        let company = sheet.get_value((2, row_idx));

        // Check if next row is empty
        if date == "" || company == "" {
            out.push(cursor);
            break;
        }

        if date == cursor.date && company == cursor.company {
            cursor.to = row_idx;
        } else {
            out.push(cursor);
            cursor = CompanyDay {
                company: company,
                date: date,
                from: row_idx,
                to: row_idx,
                max: 0,
            };
        }
    }

    for idx in 0..out.len() {
        out[idx].max = row_idx;
    }

    out
}

pub fn cut_rows(path_to_file: &std::path::PathBuf, from: u32, to: u32, max: u32) {
    let mut book = reader::xlsx::read(path_to_file).unwrap();

    if max - to + 1 > 0 {
        book.remove_row("Sheet_1", &(to + 1), &(max - to + 1));
    }

    if from > 2 {
        book.remove_row("Sheet_1", &2, &(from - 2));
    }

    let _ = writer::xlsx::write(&book, path_to_file);
}
