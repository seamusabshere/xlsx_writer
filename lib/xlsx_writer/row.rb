module XlsxWriter
  class Row
    attr_reader :sheet
    attr_reader :cells
    
    def initialize(sheet, columns)
      @sheet = sheet
      @cells = columns.map do |column|
        Cell.new self, column
      end
    end
    
    def ndx
      sheet.rows.index(self) + 1
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
