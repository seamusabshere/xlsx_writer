class XlsxWriter
  class SheetRels < Xml

    AUTO = false

    attr_reader :sheet

    def initialize(document, sheet)
      @sheet = sheet
      super document
    end
    
    def relative_path
      "xl/worksheets/_rels/sheet#{sheet.ndx}.xml.rels"
    end
  end
end
