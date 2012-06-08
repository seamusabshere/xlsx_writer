require 'bundler/setup'

if ::Bundler.definition.specs['debugger'].first
  require 'debugger'
elsif ::Bundler.definition.specs['ruby-debug'].first
  require 'ruby-debug'
end

require 'xlsx_writer'

@doc = XlsxWriter::Document.new

@sheet1 = @doc.add_sheet("Sheet1")
@sheet1.add_row(['a', 'a'])
@sheet1.add_row(['a', 'a'])
@sheet1.add_row(['a', 'a'])
# @sheet1.add_row(['foo', 'bar'])
@sheet1.add_autofilter 'A1:B1'

@sheet2 = @doc.add_sheet("Sheet2")
@sheet2.add_row(['a', 'a'])
# @sheet2.add_row(['hello', 'world'])
# @sheet2.add_row(['yo', 'there'])
# @sheet2.add_row(['foo', 'bar'])
@sheet2.add_autofilter 'A1:B1'

FileUtils.mv @doc.path, 'foo.xlsx'
@doc.cleanup
