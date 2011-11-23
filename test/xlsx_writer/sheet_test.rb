require File.dirname(__FILE__) + '/../test_helper'
require "rexml/document"
require 'date'
require 'time'

module Xlsx
  class SheetTest < Test::Unit::TestCase

    def test_column_index
      assert_equal 'A', Sheet.column_index(0)
      assert_equal 'B', Sheet.column_index(1)
      assert_equal 'C', Sheet.column_index(2)
      assert_equal 'D', Sheet.column_index(3)
      assert_equal 'Y', Sheet.column_index(24)
      assert_equal 'Z', Sheet.column_index(25)
    end

    def test_column_index_two_digits
      assert_equal 'AA', Sheet.column_index(0+26)
      assert_equal 'AB', Sheet.column_index(1+26)
      assert_equal 'AC', Sheet.column_index(2+26)
      assert_equal 'AD', Sheet.column_index(3+26)
      assert_equal 'AZ', Sheet.column_index(25+26)
      assert_equal 'BA', Sheet.column_index(25+26+1)
      assert_equal 'BB', Sheet.column_index(25+26+2)
      assert_equal 'BC', Sheet.column_index(25+26+3)
    end

    def test_format_field_for_strings
      v = Sheet.format_field_and_type_and_style :type => 'String', :value => "<escape this>"
      assert_equal [:inlineStr, "<is><t>&lt;escape this&gt;</t></is>", 3], v
    end

    def test_format_field_for_numbers
      v = Sheet.format_field_and_type_and_style :type => 'Number', :value => 3
      assert_equal [:n, "<v>3</v>", 6], v
      v = Sheet.format_field_and_type_and_style :type => 'Number', :value => 9.32
      assert_equal [:n, "<v>9.32</v>", 6], v
    end

    def test_format_field_for_date
      v = Sheet.format_field_and_type_and_style(:type => 'Date', :value => '2010-Jul-24')
      assert_equal [:n, "<v>#{38921+1462}</v>", 1], v
    end

    def test_format_field_for_datetime
      v = Sheet.format_field_and_type_and_style(:type => 'Date', :value => '2010-Jul-24 12:00 UTC')
      assert_equal [:n, "<v>#{38921+1462}</v>", 1], v
    end

    def test_format_field_for_boolean
      v = Sheet.format_field_and_type_and_style(:type => 'Boolean', :value => false)
      assert_equal [:b, "<v>0</v>", 5], v
      v = Sheet.format_field_and_type_and_style(:type => 'Boolean', :value => true)
      assert_equal [:b, "<v>1</v>", 5], v
    end

    def test_add_row
      headers = [{:type=>"Number", :width=>155, :header=>"Agreement Type"}, 
        {:type=>"String", :width=>100, :header=>"Mumboe Id"}, 
        {:type=>"String", :width=>169, :header=>"Agreement Title"}, 
        {:type=>"String", :width=>219, :header=>"Document Title"}, 
        {:type=>"String", :width=>100, :header=>"Party"}, 
        {:type=>"Date", :width=>154, :header=>"Agreement Date Created"}, 
        {:type=>"Number", :width=>228, :header=>"Security Deposit"}, 
        {:type=>"String", :width=>100, :header=>"my multi-line field"}]

      str = ""
      io = StringIO.new(str)
      Sheet.new(nil, 'name', io, headers) do |sheet|
        sheet.add_row ['this is ', 'a new row']
      end
      doc = REXML::Document.new str
      assert_equal 'worksheet', doc.root.name
      sheetdata = doc.root.elements['sheetData']
      assert sheetdata
      row = sheetdata.elements['row']
      assert row
      assert_equal '1', row.attributes['r']
      assert_equal headers.size, row.elements.to_a.size
      assert_equal [], row.elements.to_a[0].attributes.keys - ["r", "t", "s"]
    end
    
    def test_add_autofilter
      str = ""
      io = StringIO.new(str)
      Sheet.new(nil, 'mysheet', io) do |sheet|
        sheet.add_autofilter('A1:B1')
        sheet.add_row [1, 2]
      end
      assert str.include?('</sheetData><autoFilter ref="A1:B1"/></worksheet>')
    end
  end
end
