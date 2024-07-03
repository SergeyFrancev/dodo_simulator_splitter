use std::fs;
use edit_xlsx::{Read, WorkSheetCol, WorkSheetRow, Workbook, WorkbookResult, Write};

// static SHEET_NAME: str = "Sheet_1";

fn get_path_by_name(file_name: &str) -> (String , String) {
    let res: Vec<&str> = file_name.split("__").collect();
    let cat = res[0];
    let file_date = res[1];
    (format!("{}/data/{}__{}", env!("CARGO_MANIFEST_DIR"), cat, file_date), cat.to_string())
}

fn split_file(file_name: String) -> WorkbookResult<()> {
    let (read_path, read_cat) = get_path_by_name(&file_name);
    // Read an existed workbook
    let reading_book = Workbook::from_path(read_path)?;
    let reading_sheet = reading_book.get_worksheet_by_name("Sheet_1")?;
    // Create a new workbook to write
    let mut writing_book = Workbook::new();
    let writing_sheet = writing_book.add_worksheet_by_name("Sheet_1")?;

    // Synchronous column width and format
    let columns_map = reading_sheet.get_columns_with_format("A:W")?;
    match reading_sheet.get_default_column() {
        None => writing_sheet.set_default_column_adaptive(),
        Some(width) => writing_sheet.set_default_column(width),
    }
    columns_map
        .iter()
        .for_each(|(col_range, (column, format))| {
            if let Some(format) = format {
                // if col format exists, write it to writing_sheet
                writing_sheet
                    .set_columns_with_format(col_range, column, format)
                    .unwrap()
            } else {
                writing_sheet.set_columns(col_range, column).unwrap()
            }
        });

    // Synchronous row height and format
    writing_sheet.set_default_row(reading_sheet.get_default_row());
    for row_number in 1..=reading_sheet.max_row() {
        let (row, format) = reading_sheet.get_row_with_format(row_number)?;
        if let Some(format) = format {
            // if col format exists, write it to writing_sheet
            writing_sheet.set_row_with_format(row_number, &row, &format)?;
        } else {
            writing_sheet.set_row(row_number, &row)?;
        }
    }

    // Read then write text and format
    for row in 1..=reading_sheet.max_row() {
        for col in 1..=reading_sheet.max_column() {
            if let Ok(cell) = reading_sheet.read_cell((row, col)) {
                writing_sheet.write_cell((row, col), &cell)?;
            }
        }
    }

    let path = format!("{}/data/{}__new.xlsx", env!("CARGO_MANIFEST_DIR"), read_cat);
    fs::remove_file(&path);
    writing_book.save_as(path)?;
    Ok(())
}

fn main() {
    split_file("Метрики_заказа__фвыапф.xlsx".to_string());
    split_file("Смены_сотрудников__фывафу.xlsx".to_string());
    split_file("Состав_заказа__фывафыв.xlsx".to_string());
}

// fn main() {
//     // println!("Hello, world!");

//     let path = format!("{}/data/Метрики_заказа__фвыапф.xlsx", env!("CARGO_MANIFEST_DIR"));
//     let mut workbook: Xlsx<_> = open_workbook(path).unwrap();
//     let range = workbook
//         .worksheet_range("Sheet_1")
//         .unwrap();
//         // .ok_or(Error::Msg("Cannot find 'Sheet_1'"))??;

//     let mut iter = RangeDeserializerBuilder::new().from_range(&range).unwrap();

//     if let Some(result) = iter.next() {
//         let (label, value): (String, String) = result.unwrap();
//         println!("label: {}, value: {}", label, value);
//         // assert_eq!(label, "celsius");
//         // assert_eq!(value, 22.2222);
//         // Ok(())
//     } else {
//         // Err(From::from("expected at least one record but got none"))
//         println!("ERRRRRR");
//     }
// }
