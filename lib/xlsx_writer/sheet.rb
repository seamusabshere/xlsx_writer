require 'fast_xs'

class XlsxWriter
  class Sheet
    class << self
      def excel_name(value)
        str = value.to_s.dup
        str.gsub! '/', ''   # remove forward slashes
        str.gsub! /\s+/, '' # compress "inner" whitespace
        str.strip!          # trim whitespace from ends
        str.fast_xs
      end
    end

    BUFSIZE = 131072 #128kb

    attr_reader :document
    attr_reader :name
    attr_reader :ndx
    attr_reader :autofilters
    attr_reader :path
    attr_reader :row_count
    attr_reader :max_row_length
    attr_reader :max_cell_pixel_width

    # Freeze the pane under this top left cell
    attr_accessor :freeze_top_left

    def initialize(document, name, ndx)
      @mutex = Mutex.new
      @document = document
      @ndx = ndx
      @name = Sheet.excel_name name
      @row_count = 0
      @autofilters = []
      @max_row_length = 1
      @max_cell_pixel_width = Hash.new(Cell.pixel_width(5))
      @path = ::File.join document.staging_dir, relative_path
      ::FileUtils.mkdir_p ::File.dirname(path)
      @rows_tmp_file_writer = ::File.open(rows_tmp_file_path, 'wb')
    end

    def generated?
      @generated == true
    end

    def generate
      return if generated?
      @mutex.synchronize do
        return if generated?
        @generated = true
        @rows_tmp_file_writer.close
        File.open(path, 'wb') do |f|
          f.write <<-EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
EOS
          if freeze_top_left
            f.write <<-EOS
<sheetViews>
  <sheetView workbookViewId="0">
    <pane ySplit="#{y_split}" topLeftCell="#{freeze_top_left}" activePane="topLeft" state="frozen"/>
  </sheetView>
</sheetViews>
EOS
          end
          f.write %{<cols>}
          (0..max_row_length-1).each do |x|
            f.write %{<col min="#{x+1}" max="#{x+1}" width="#{max_cell_pixel_width[x]}" bestFit="1" customWidth="1" />}
          end
          f.write %{</cols>}
          f.write %{<sheetData>}
          File.open(rows_tmp_file_path, 'rb') do |rows_tmp_file_reader|
            buffer = ''
            while rows_tmp_file_reader.read(BUFSIZE, buffer)
              f.write buffer
            end
          end
          f.write %{</sheetData>}
          autofilters.each { |autofilter| f.write autofilter.to_xml }
          f.write document.page_setup.to_xml
          f.write document.header_footer.to_xml
          f.write %{</worksheet>}
        end
        File.unlink rows_tmp_file_path
        converted = UnixUtils.unix2dos path
        FileUtils.mv converted, path
        SheetRels.new(document, self).generate
      end
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
        
    def add_row(cells)
      raise ::RuntimeError, "Can't add row, already generated!" if generated?
      @row_count += 1
      row = Row.new self, cells, row_count
      @rows_tmp_file_writer.write row.to_xml
      if (l = row.cells.length) > max_row_length
        @max_row_length = l
      end
      row.cells.each_with_index do |cell, x|
        if (w = cell.pixel_width) > max_cell_pixel_width[x]
          max_cell_pixel_width[x] = w
        end
      end
      nil
    end

    private

    def rows_tmp_file_path
      path + '.rows_tmp_file'
    end

    def y_split
      if freeze_top_left =~ /(\d+)$/
        $1.to_i - 1
      else
        raise "freeze_top_left must be like 'A3', was #{freeze_top_left}"
      end
    end
  end
end
