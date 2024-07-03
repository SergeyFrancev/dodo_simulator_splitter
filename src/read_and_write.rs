use calamine::{
    open_workbook, open_workbook_auto, Data, DataType, Error, RangeDeserializerBuilder, Reader,
    Xlsx,
};
use chrono::{Datelike, Duration, NaiveDate, NaiveDateTime, Timelike};
use std::fs;
// use chrono::Datelike;
use xlsxwriter::{worksheet::DateTime, Format, Workbook};

fn write() {
    let path = format!("{}/data/out.xlsx", env!("CARGO_MANIFEST_DIR"));
    let workbook = Workbook::new(&path).unwrap();

    let mut sheet1 = workbook.add_worksheet(None).unwrap();
    sheet1.write_string(0, 0, "Red text", None).unwrap();
    // sheet1.write_number(0, 1, 20., None)?;
    // sheet1.write_formula_num(1, 0, "=10+B1", None, 30.)?;
    // sheet1.write_url(
    //     1,
    //     1,
    //     "https://github.com/informationsea/xlsxwriter-rs",
    //     Some(&Format::new().set_font_color(FormatColor::Blue).set_underline(FormatUnderline::Single)),
    // )?;
    // sheet1.merge_range(2, 0, 3, 2, "Hello, world", Some(
    //     &Format::new().set_font_color(FormatColor::Green).set_align(FormatAlignment::CenterAcross)
    //                   .set_vertical_align(FormatVerticalAlignment::VerticalCenter)))?;

    // sheet1.set_selection(1, 0, 1, 2);
    // sheet1.set_tab_color(FormatColor::Cyan);
    workbook.close().unwrap();
}

fn excel_date_to_naive_date(excel_date: f64) -> NaiveDate {
    let base_date = NaiveDate::from_ymd(1899, 12, 30); // Excel considers 1900-01-01 as day 1, so we use 1899-12-30 as the base
    let days = excel_date.trunc() as i64;
    base_date + Duration::days(days)
}

fn excel_date_to_naive_datetime(excel_date: f64) -> NaiveDateTime {
    let date_part = excel_date_to_naive_date(excel_date);
    let time_part = (excel_date.fract() * 86400.0) as i64; // 86400 seconds in a day
    date_part.and_hms(0, 0, 0) + Duration::seconds(time_part)
}

fn naive_datetime_to_xlsx_datetime(naive_datetime: NaiveDateTime) -> DateTime {
    DateTime::new(
        naive_datetime.year() as i16,
        naive_datetime.month() as i8,
        naive_datetime.day() as i8,
        naive_datetime.hour() as i8,
        naive_datetime.minute() as i8,
        naive_datetime.second() as f64,
    )
}

fn split_file(file_name: String) {
    let res: Vec<&str> = file_name.split("__").collect();
    let cat = res[0];
    let file_date = res[1];

    // println!("cat: {}, file_date: {}", cat, file_date);
    let path = format!(
        "{}/data/{}__{}",
        env!("CARGO_MANIFEST_DIR"),
        cat,
        file_date
    );
    let mut workbook = open_workbook_auto(path).expect("Cannot open file");
    let range = workbook.worksheet_range("Sheet_1").unwrap();
    // let mut iter = RangeDeserializerBuilder::new().from_range(&range).unwrap();

    let path = format!("{}/data/{}__new.xlsx", env!("CARGO_MANIFEST_DIR"), cat);
    fs::remove_file(path.clone());
    let workbook_out = Workbook::new(&path).unwrap();
    let mut sheet1 = workbook_out.add_worksheet(None).unwrap();

    // Выберите активный лист
    // if let Some(Ok(range)) = workbook.worksheet_range("Sheet_1") {
    let mut row_idx = 0;
    let mut prev_date: Option<f64> = None;
    let mut prev_point: Option<String> = None;
    let mut quit = false;
    for row in range.rows() {
        if row_idx == 0 {
            // TODO: store header

            let mut col_idx = 0;
            for cell in row {
                sheet1
                    .write_string(row_idx, col_idx, &cell.to_string(), None)
                    .unwrap();
                col_idx += 1;
            }
            row_idx += 1;
            continue;
        }
        let mut col_idx = 0;
        for cell in row {
            match cell {
                Data::DateTime(date) => {
                    let d = excel_date_to_naive_datetime(date.as_f64());
                    let nd = naive_datetime_to_xlsx_datetime(d);

                    if col_idx == 0 {
                        if prev_date == None {
                            prev_date = Some(date.as_f64());
                        }
                        if prev_date != Some(date.as_f64()) {
                            quit = true;
                            break;
                        }
                    }

                    let mut date_format = "M/d/yyyy h:mm:ss";
                    if nd.day == 30 && nd.month == 12 && nd.year == 1899 {
                        date_format = "h:mm:ss";
                        let h = nd.hour.to_string();
                        let m = if nd.min < 10 {
                            format!("0{}", nd.min)
                        } else {
                            nd.min.to_string()
                        };
                        let s = if nd.second < 10f64 {
                            format!("0{}", nd.second)
                        } else {
                            (nd.second as i64).to_string()
                        };
                        sheet1
                            .write_string(row_idx, col_idx, &format!("{}:{}:{}", h, m, s), None)
                            .unwrap();
                    } else {
                        let mut date_format = "M/d/yyyy h:mm:ss";
                        if nd.hour == 0 && nd.min == 0 && nd.second == 0f64 {
                            date_format = "M/d/yyyy";
                        } else {
                            date_format = "M/d/yyyy h:mm:ss";
                        }

                        sheet1.write_datetime(row_idx, col_idx, &nd, None).unwrap();
                        sheet1
                            .set_column(
                                col_idx,
                                col_idx,
                                20.0,
                                Some(&Format::new().set_num_format(date_format)),
                            )
                            .unwrap();
                    }
                }
                _ => {
                    if col_idx == 1 {
                        if prev_point == None {
                            prev_point = Some(cell.to_string());
                        }
                        if prev_point != Some(cell.to_string()) {
                            quit = true;
                            break;
                        }
                    }
                    sheet1
                        .write_string(row_idx, col_idx, &cell.to_string(), None)
                        .unwrap();
                }
            }

            col_idx += 1;
        }
        if quit {
            break;
        }
        row_idx += 1;
    }
    workbook_out.close().unwrap();
}

fn main() {
    split_file("Метрики_заказа__фвыапф.xlsx".to_string());
    split_file("Смены_сотрудников__фывафу.xlsx".to_string());
    split_file("Состав_заказа__фывафыв.xlsx".to_string());
    // write();
    // }
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
