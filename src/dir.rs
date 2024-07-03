use std::{env, fs, path};
use chrono::{NaiveDate, Duration};

fn excel_date_to_naive_date(excel_date: f64) -> NaiveDate {
    let base_date = NaiveDate::from_ymd(1899, 12, 30); // Excel considers 1900-01-01 as day 1, so we use 1899-12-30 as the base
    let days = excel_date.trunc() as i64;
    base_date + Duration::days(days)
}

pub fn get_executable_directory() -> Result<String, String> {
    // Получение полного пути к исполняемому файлу
    let exe_path =
        env::current_exe().map_err(|e| format!("Failed to get the executable path: {}", e))?;

    // Извлечение директории из полного пути
    let exe_dir = exe_path
        .parent()
        .ok_or("Failed to get the executable directory".to_string())?;

    // Возвращаем директорию как PathBuf
    Ok(exe_dir.to_str().to_owned().unwrap().to_string())
}

pub fn get_sources() -> Result<Vec<String>, String> {
    // Specify the directory path
    let dir_path = get_executable_directory()?; // current directory

    // Read the directory
    let entries = fs::read_dir(dir_path).expect("Failed to read the directory");

    let mut out: Vec<String> = Vec::new();
    // Iterate over the entries and print their names
    for entry in entries {
        let entry = entry.expect("Inavalid entry");
        let path = entry.path();
        let src_files = Vec::from(["Метрики_заказа", "Смены_сотрудников", "Состав_заказа"]);

        // Get the file name as a string
        if let Some(file_name) = path.file_name() {
            for src_file in src_files.iter() { 
                if file_name.to_string_lossy().to_string().starts_with(src_file) {
                    out.push(file_name.to_string_lossy().to_string());
                }   
            }
        }
    }

    Ok(out)
}

fn result_dir() -> path::PathBuf {
    let dir = get_executable_directory().unwrap();
    path::PathBuf::from(dir.as_str()).join("result")
}

pub fn recreate_result_dir() -> path::PathBuf {
    let dir_path = result_dir();
    if dir_path.exists() {
        fs::remove_dir_all(&dir_path).unwrap();
    }
    fs::create_dir_all(&dir_path).unwrap();
    dir_path
}

pub fn create_part(src_file: &path::PathBuf, company: String, date: String) -> path::PathBuf {
    let mut dir_path = result_dir().join(company);

    // Create company dir if not exist
    if !dir_path.exists() {
        fs::create_dir_all(&dir_path).unwrap();
    }

    // Formate date
    let date_val = date.parse::<f64>().unwrap();
    let date_date = excel_date_to_naive_date(date_val);
    dir_path = dir_path.join(date_date.format("%Y-%m-%d").to_string());

    // Create date dir if not exist
    if !dir_path.exists() {
        fs::create_dir_all(&dir_path).unwrap();
    }

    // Copy file
    let file_name = src_file.file_name().unwrap().to_str().unwrap();
    dir_path = dir_path.join(file_name);
    fs::copy(src_file.as_path(), dir_path.as_path()).unwrap();

    dir_path
}