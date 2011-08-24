require 'bigdecimal'
require 'time'

module SimpleXlsx
  class Sheet
    attr_reader :name
    attr_accessor :rid

    #column information is an array of hashes containing information about the columns, this is used to figure out the width of each column
    #send it something like
    # [{:type => "String", :value => "First Column", :width => 100},{:type => "Number", :value => "Second Column", :width => 200}]
    # it will use this information to create the header row and setup the widths for the columns
    def initialize document, name, column_information, stream, &block
      @document = document
      @stream = stream
      @name = Sheet::xscape(name)
      @row_ndx = 1
      @stream.write <<-ends.lf_to_crlf
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<dimension ref="A1:F10"/>
<sheetViews>
  <sheetView tabSelected="1" workbookViewId="0">
    <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
    <selection pane="bottomLeft"/>
  </sheetView>
</sheetViews>
<sheetFormatPr defaultRowHeight="15"/>
ends
      if Sheet::blank?(column_information)
        @stream.write("<sheetData>".lf_to_crlf)
      else
        self.write_column_header(column_information)
      end
      if block_given?
        yield self
      end
      @stream.write "</sheetData></worksheet>".lf_to_crlf
    end

    # width = ( (pixel_width + 5)/(8*256))*256
    # the + 5 is the padding pixels put in by MS, 8 is the pixel width of the font (best guess)
    def write_column_header(column_info)
      num_cols = column_info.size
      @stream.write("<cols>".lf_to_crlf)
      column_info.each_with_index do |c, i|
        width = ((c[:width] + 5)/2048.to_f)*256
        @stream.write("<col min=\"#{i+1}\" max=\"#{i+1}\" width=\"#{width.to_s}\" bestFit=\"1\" customWidth=\"1\"/>".lf_to_crlf)
      end
      @stream.write("</cols>".lf_to_crlf)
      @stream.write("<sheetData>".lf_to_crlf)
      self.add_row column_info, true
    end
    
    # for more control over styling, pass in array of hash values, for example
    # [{:type => "DateTime", :width => 100, :value => "Date"},{:type => "String", :width => 800, :value => "Long String"}]
    def add_row arry, header = false
      if header
        row = ["<row r=\"#{@row_ndx}\" s=\"4\" customFormat=\"1\">"]
        cstyle = 4
        kind = :inlineStr
      else
        row = ["<row r=\"#{@row_ndx}\">"]
      end
      arry.each_with_index do |data_hash, col_ndx|
        if header
          ccontent = "<is><t>#{xscape(data_hash[:value])}</t></is>"
        else
          kind, ccontent, cstyle = Sheet.format_field_and_type_and_style data_hash
        end
        t_kind = blank?(kind) ? "" : "t=\"#{kind.to_s}\""
        t_style = blank?(cstyle) ? "" : "s=\"#{cstyle}\""
        if blank?(ccontent)
          row << "<c r=\"#{Sheet.column_index(col_ndx)}#{@row_ndx}\" #{t_kind} #{t_style}/>"
        else
          row << "<c r=\"#{Sheet.column_index(col_ndx)}#{@row_ndx}\" #{t_kind} #{t_style}>#{ccontent}</c>"
        end
      end
      row << "</row>"
      @row_ndx += 1
      @stream.write(row.join("\r\n"))
    end

    def xscape(value)
      Sheet::xscape(value)
    end

    def blank?(object)
      Sheet::blank?(object)
    end

    def self.format_field_and_type_and_style data_hash
      if data_hash[:type] == "String"
        if blank?(data_hash[:value])
          [:inlineStr, "", 3]
        else
          [:inlineStr, "<is><t>#{xscape(data_hash[:value])}</t></is>", 3]
        end
      elsif data_hash[:type] == "Number"
        if Sheet.is_multilined?(data_hash[:value])
          [:inlineStr, "<is><t>#{xscape(data_hash[:value])}</t></is>", 6]
        else
          [:n, "<v>#{xscape(data_hash[:value])}</v>", 6]
        end
      elsif data_hash[:type] == "DateTime"
        if blank?(data_hash[:value])
          [:inlineStr, "", 3]
        else
          if Sheet.is_multilined?(data_hash[:value])
            [:inlineStr, "<is><t xml:space=\"preserve\">#{data_hash[:value]}</t></is>", 1]
          else
            [:n, "<v>#{days_since_jan_1_1900(Date.parse(data_hash[:value]))}</v>", 1]
          end
        end
      elsif data_hash[:type] == "Boolean"
        if Sheet.is_multilined?(data_hash[:value])
          [:inlineStr, "<is><t>#{data_hash[:value].upcase}</t></is>", 5]
        else
          [:b, "<v>#{data_hash[:value] ? '1' : '0'}</v>", 5]
        end
      elsif data_hash[:type] == "Money"
        if blank?(data_hash[:value])
          [:n, "<v>#{data_hash[:value]}</v>", 2]
        else
          if Sheet.is_multilined?(data_hash[:value])
            [:inlineStr, "<is><t>#{xscape(data_hash[:value])}</t></is>", 2]
          else
            [:n, "<v>#{xscape(data_hash[:value])}</v>", 2]
          end
        end
      else
        [:inlineStr, "<is><t>#{xscape(data_hash[:value])}</t></is>", 3]
      end
    end

    def self.days_since_jan_1_1900 date
      @@jan_1_1904 ||= Date.parse("1904 Jan 1")
      (date - @@jan_1_1904).to_i + 1462 # http://support.microsoft.com/kb/180162
    end

    def self.fractional_days_since_jan_1_1900 value
      @@jan_1_1904_midnight ||= ::Time.utc(1904, 1, 1)
      ((value - @@jan_1_1904_midnight) / 86400.0) + #24*60*60
        1462 # http://support.microsoft.com/kb/180162
    end

    def self.abc
      @@abc ||= ('A'..'Z').to_a
    end

    def self.column_index n
      result = []
      while n >= 26 do
        result << abc[n % 26]
        n /= 26
      end
      result << abc[result.empty? ? n : n - 1]
      result.reverse.join
    end
    
    # use this to sub out values that excel doesn't like, for example & changing to &amp;
    def self.clean_string(value)
      value.gsub(/\&/, '&amp;')
    end
    
    def self.clean_number(value)
      value.gsub!(/\$/, '')
      value.gsub(/\,/, '')
    end
    
    def self.is_multilined?(value)
      value=~/\r|\n/ ? true : false
    end

    # add a Rails-y blank? for convenience
    def self.blank?(object)
      object.respond_to?(:empty?) ? object.empty? : !object
    end

    def self.xscape(value)
      value.respond_to?(:to_xs) ? value.to_xs : value
    end
  end
end
