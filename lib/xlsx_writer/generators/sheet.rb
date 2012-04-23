require 'fast_xs'

module XlsxWriter
  class Sheet < Xml
    class << self
      def excel_name(value)
        str = value.to_s.dup
        str.gsub! '/', ''   # remove forward slashes
        str.gsub! /\s+/, '' # compress "inner" whitespace
        str.strip!          # trim whitespace from ends
        str.fast_xs
      end
    end

    AUTO = false
    
    attr_reader :name
    attr_reader :rows
    attr_reader :autofilters

    def initialize(document, name)
      @name = Sheet.excel_name name
      @rows = []
      @autofilters = []
      super document
    end
    
    def ndx
      document.sheets.index(self) + 1
    end

    # +1 because styles.xml occupies the first spot
    def rid
      "rId#{ndx + 1}"
    end
    
    def relative_path
      "xl/worksheets/sheet#{ndx}.xml"
    end
    
    def absolute_path
      "/#{relative_path}"
    end

    # specify range like "A1:C1"
    def add_autofilter(range)
      raise ::RuntimeError, "Can't add autofilter, already generated!" if generated?
      autofilters << Autofilter.new(range)
    end
        
    def add_row(data)
      raise ::RuntimeError, "Can't add row, already generated!" if generated?
      row = Row.new self, data
      rows << row
      row
    end
        
    # override Xml method to save memory
    def path
      @path || @mutex.synchronize do
        @path ||= begin
          memo = ::File.join document.staging_dir, relative_path
          ::FileUtils.mkdir_p ::File.dirname(memo)
          ::File.open(memo, 'wb') do |f|
            to_file f
          end
          converted = UnixUtils.unix2dos memo
          ::FileUtils.mv converted, memo
          SheetRels.new(document, self).path
          @generated = true
          memo
        end
      end
    end

    private
    
    # not using ERB to save memory
    def to_file(f)
      f.puts <<-EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<cols>
EOS
      (0..max_length-1).each do |x|
        f.puts %{<col min="#{x+1}" max="#{x+1}" width="#{max_cell_width(x)}" bestFit="1" customWidth="1" />}
      end
      f.puts %{</cols>}
      f.puts %{<sheetData>}
      rows.each { |row| f.puts row.to_xml }
      f.puts %{</sheetData>}
      autofilters.each { |autofilter| f.puts autofilter.to_xml }
      f.puts document.page_setup.to_xml
      f.puts document.header_footer.to_xml
      f.puts %{</worksheet>}
    end
    
    def max_length
      if max = rows.max_by { |row| row.length }
        max.length
      else
        1
      end
    end
    
    def max_cell_width(x)
      if max = rows.max_by { |row| row.cell_width(x) }
        max.cell_width x
      else
        Cell.pixel_width 5
      end
    end
  end
end
