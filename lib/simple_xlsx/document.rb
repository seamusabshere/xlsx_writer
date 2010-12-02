module SimpleXlsx
  class Document
    def initialize(io)
      @sheets = []
      @io = io
    end

    attr_reader :sheets

    def add_sheet name, &block
      # Naview: block not required as existing stream is returned
      stream = @io.open_stream_for_sheet(@sheets.size)
#      @io.open_stream_for_sheet(@sheets.size) do |stream|
        @sheets << Sheet.new(self, name, stream, &block)
#      end
    end

    def has_shared_strings?
      false
    end

  end
end
