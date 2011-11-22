# xlsx_writer

Writes (doesn't read) XLSX files.

## Credit

https://github.com/harvesthq/simple_xlsx_writer

https://github.com/mumboe/simple_xlsx_writer

## Example

    require 'xlsx_writer'

    doc = XlsxWriter::Document.new

    sheet1 = doc.add_sheet("People")

    sheet1.add_autofilter "A1:B1"
    sheet1.add_row(%w{DoB Name Occupation})
    sheet1.add_row([Date.parse("July 31, 1912"), 
                   "Milton Friedman", 
                   "Economist / Statistician"])

    FileUtils.mv doc.path, 'people.xlsx'

    doc.cleanup
