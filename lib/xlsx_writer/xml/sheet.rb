require 'fast_xs'

class XlsxWriter
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

    # Freeze the pane under this top left cell
    attr_accessor :freeze_top_left

    def initialize(document, name)
      @name = Sheet.excel_name name
      @rows = []
      @autofilters = []
      super document
    end
    
    def ndx
      document.sheets.index(self) + 1
    end

    def local_id
      ndx - 1
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
      autofilters << Autofilter.new(self, range)
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
    
    def y_split
      if freeze_top_left =~ /(\d+)$/
        $1.to_i - 1
      else
        raise "freeze_top_left must be like 'A3', was #{freeze_top_left}"
      end
    end

    # not using ERB to save memory
    def to_file(f)
      f.write <<-EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
EOS
      if freeze_top_left
        f.write <<-EOS
<sheetViews>
  <sheetView workbookViewId="0">
    <pane ySplit="#{y_split}" topLeftCell="#{freeze_top_left}" activePane="bottomLeft" state="frozen"/>
  </sheetView>
</sheetViews>
EOS
      end
      f.write %{<cols>}
      (0..max_length-1).each do |x|
        f.write %{<col min="#{x+1}" max="#{x+1}" width="#{max_cell_width(x)}" bestFit="1" customWidth="1" />}
      end
      f.write %{</cols>}
      f.write %{<sheetData>}
      rows.each { |row| f.write row.to_xml }
      f.write %{</sheetData>}
      autofilters.each { |autofilter| f.write autofilter.to_xml }
      f.write document.page_setup.to_xml
      f.write document.header_footer.to_xml
      f.write %{</worksheet>}
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
