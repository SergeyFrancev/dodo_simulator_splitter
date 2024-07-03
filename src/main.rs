use dodo_simulator_splitter::{get_executable_directory, get_ranges, get_sources};
use std::path;

fn main() {
    // Remove and recreate 'result' dir
    let _ = dodo_simulator_splitter::recreate_result_dir();

    // Get three files
    let src_dir = get_executable_directory().expect("Invalid source directory");
    let sources = get_sources().expect("Invalid source files");

    // Read every file and copy to 'result' dir
    for source in sources {
        // Calc range for every file
        let file_path = path::PathBuf::from(&src_dir).join(&source);
        println!("*{}", file_path.display());
        let ranges = get_ranges(&file_path);

        // Create part for every range
        for range in ranges {
            println!("{}:{} | {}-{} [{}]", range.date, range.company, range.from, range.to, range.max);
            let part_path = dodo_simulator_splitter::create_part(&file_path, range.company, range.date);

            dodo_simulator_splitter::cut_rows(&part_path, range.from, range.to, range.max);
        }
    }
}
