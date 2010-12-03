require 'test_helper.rb'

module SimpleXlsx
  class DocumentTest < Test::Unit::TestCase
    def open_stream_for_sheet sheets_size
      self
    end

    def write arg
      # This space intentionally left blank
    end

    def test_add_sheet
      @doc = Document.new self
      assert_equal [], @doc.sheets
      @doc.add_sheet "new sheet"
      assert_equal 1, @doc.sheets.size
      assert_equal 'new sheet', @doc.sheets.first.name
    end
    
    def test_add_sheet_with_forward_slash_in_name
      @doc = Document.new self
      assert_equal [], @doc.sheets
      @doc.add_sheet "new sheet with /"
      assert_equal 1, @doc.sheets.size
      assert_equal 'new sheet with', @doc.sheets.first.name
    end
  end
end
