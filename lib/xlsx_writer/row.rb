class XlsxWriter
  class Row
    attr_reader :sheet
    attr_reader :cells
    attr_reader :ndx
    attr_reader :width
    
    def initialize(sheet, cells, ndx)
      @sheet = sheet
      @ndx = ndx
      @cells = cells.map do |cell|
        Cell.new self, cell
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
