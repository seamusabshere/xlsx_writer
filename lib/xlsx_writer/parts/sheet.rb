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
    
    attr_reader :name

    def initialize(document, name)
      @document = document
      @name = Sheet.excel_name name
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

    # specify range like "A1:C1"
    def add_autofilter(range)
      raise ::RuntimeError, "Can't add autofilter, already generated!" if generated?
      autofilters << range
    end
    
    def autofilters
      @autofilters ||= []
    end
    
    def rows
      @rows ||= []
    end
    
    # for more control over styling, pass in array of hash values, for example
    # [{:type => "DateTime", :width => 100, :value => "Date"},{:type => "String", :width => 800, :value => "Long String"}]
    def add_row(data)
      raise ::RuntimeError, "Can't add row, already generated!" if generated?
      row = Row.new self, data
      rows << row
      row
    end
  end
end
