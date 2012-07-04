#!/usr/bin/env ruby
require 'bundler/setup'

if ::Bundler.definition.specs['debugger'].first
  require 'debugger'
elsif ::Bundler.definition.specs['ruby-debug'].first
  require 'ruby-debug'
end

require 'xlsx_writer'

@doc = XlsxWriter::Document.new

# @sheet1 = @doc.add_sheet("Sheet1")
# @sheet1.add_row(['a', 'a'])
# @sheet1.add_row(['a', { :value => 'a', :faded => true, :type => :String }])
# @sheet1.add_row(['a', 'a'])
# # @sheet1.add_row(['foo', 'bar'])
# @sheet1.add_autofilter 'A1:B1'

@sheet2 = @doc.add_sheet("Sheet2")
@sheet2.add_row(['a', 'a'])
@sheet2.add_row(['false1', false])
@sheet2.add_row(['false2', {:value => false, :type => :Boolean}])
@sheet2.add_row(['false3', 'faLse'])
@sheet2.add_row(['true1', true])
@sheet2.add_row(['true2', {:value => true, :type => :Boolean}])
@sheet2.add_row(['true3', 'trUe'])

# @sheet2.add_row(['hello', 'world'])
# @sheet2.add_row(['yo', 'there'])
# @sheet2.add_row(['foo', 'bar'])
@sheet2.add_autofilter 'A1:B1'

FileUtils.mv @doc.path, 'foo.xlsx'
@doc.cleanup
