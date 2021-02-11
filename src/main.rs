use calamine::{open_workbook, DataType, Reader, Xlsx};

fn main() {
  let path = "ERR_AG0047_20210209_00005479.xlsx";
  let mut wtr = csv::WriterBuilder::new()
    .delimiter(b';')
    .from_path("teste4.csv")
    .expect("msg: &str");

  let mut workbook: Xlsx<_> = open_workbook(path).expect("Cannot open file!");

  let range = workbook
    .worksheet_range("Sheet1")
    .unwrap()
    .expect("msg: &str");

  let rows = range.rows();

  for (_i, row) in rows.enumerate() {
    let cols: Vec<String> = row
      .iter()
      .map(|c| match c {
        DataType::String(c) => format!("{}", c),
        _ => "".to_string(),
      })
      .collect();
    wtr.write_record(&cols).unwrap();
  }
  wtr.flush().unwrap();
}
