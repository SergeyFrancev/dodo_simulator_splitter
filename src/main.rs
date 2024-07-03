use dodo_simulator_splitter::{
    create_part, cut_rows, get_executable_directory, get_ranges, get_sources, CompanyDay,
};
use rayon::prelude::*;

fn process_path(src_file: &std::path::PathBuf, range: &CompanyDay) {
    println!(
        "{}:{} | {}-{} [{}]",
        range.date, range.company, range.from, range.to, range.max
    );
    let part_path = create_part(&src_file, range.company.clone(), range.date.clone());

    cut_rows(&part_path, range.from, range.to, range.max);
}

fn process_file(path: &String) {
    let src_dir = get_executable_directory().expect("Invalid source directory");
    let file_path = src_dir.join(&path);
    println!("*{}", file_path.display());

    // Calc range for every file
    get_ranges(&file_path)
        .par_iter()
        .for_each(|range| process_path(&file_path, range));
}

fn main() {
    get_sources()
        .expect("Invalid source files")
        .par_iter()
        .for_each(process_file);
}
