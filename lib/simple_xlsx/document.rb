module SimpleXlsx
  class Document
    def initialize(io)
      @sheets = []
      @io = io
    end

    attr_reader :sheets

    def add_sheet name, &block
      stream = @io.open_stream_for_sheet(@sheets.size)
      @sheets << Sheet.new(self, escape_for_excel(name), stream, &block)
    end

    def has_shared_strings?
      false
    end
    
    def escape_for_excel(name)
      name.gsub(/\//, "").strip # Remove forward slashes and trim white space
    end
  end
end
