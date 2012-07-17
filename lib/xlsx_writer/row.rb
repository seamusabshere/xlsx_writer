class XlsxWriter
  class Row
    attr_reader :sheet
    attr_reader :cells
    attr_reader :width
    
    def initialize(sheet, columns)
      @width = {}
      @sheet = sheet
      @cells = columns.map do |column|
        Cell.new self, column
      end
    end
    
    def ndx
      sheet.rows.index(self) + 1
    end
    
    def length
      cells.length
    end
    
    def cell_width(x)
      @width[x] ||= if (cell = cells[x])
        cell.pixel_width
      else
        0
      end
    end
    
    def to_xml
      ary = []
      ary << %{<row r="#{ndx}">}
      cells.each do |cell|
        ary << cell.to_xml
      end
      ary << %{</row>}
      ary.join
    end
  end
end
