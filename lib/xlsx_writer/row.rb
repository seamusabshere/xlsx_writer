class XlsxWriter
  class Row
    attr_reader :sheet
    attr_reader :cells
    attr_reader :y
    
    def initialize(sheet, raw_cells, y)
      @sheet = sheet
      @y = y
      @cells = []
      raw_cells.each_with_index do |cell, x|
        @cells << Cell.new(self, cell, x, y)
      end
    end
    
    def to_xml
      ary = []
      ary << %{<row r="#{y}">}
      cells.each do |cell|
        ary << cell.to_xml
      end
      ary << %{</row>}
      ary.join
    end
  end
end
