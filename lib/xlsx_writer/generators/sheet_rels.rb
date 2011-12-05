module XlsxWriter
  class SheetRels < Xml

    AUTO = false

    attr_reader :sheet

    def initialize(document, sheet)
      @document = document
      @sheet = sheet
    end
    
    def relative_path
      "xl/worksheets/_rels/sheet#{sheet.ndx}.xml.rels"
    end
  end
end
