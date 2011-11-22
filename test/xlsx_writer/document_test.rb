require File.dirname(__FILE__) + '/../test_helper'

module Xlsx
  class DocumentTest < Test::Unit::TestCase

    # mocking IO methods so test can be used as Document.new parameter
    def open_stream_for_sheet sheets_size
      self
    end

    def write arg
    end

    def close
    end


    def setup
      @headers = [{:type=>"Number", :width=>155, :header=>"Agreement Type"}, 
        {:type=>"String", :width=>100, :header=>"Mumboe Id"}, 
        {:type=>"String", :width=>169, :header=>"Agreement Title"}, 
        {:type=>"String", :width=>219, :header=>"Document Title"}, 
        {:type=>"String", :width=>100, :header=>"Party"}, 
        {:type=>"Date", :width=>154, :header=>"Agreement Date Created"}, 
        {:type=>"Number", :width=>228, :header=>"Security Deposit"}, 
        {:type=>"String", :width=>100, :header=>"my multi-line field"}]

      @rows = [{:type=>"Number", :width=>155, :value => "1234"}, 
        {:type=>"String", :width=>100, :value => "JMUHQT"}, 
        {:type=>"String", :width=>169, :value=>"test agreement title here"}, 
        {:type=>"String", :width=>219, :value=>"test document title here"}, 
        {:type=>"String", :width=>100, :value => "test party here"}, 
        {:type=>"Date", :width=>154, :value=>"test agreement created date here"}, 
        {:type=>"Number", :width=>228, :value=>"1000"}, 
        {:type=>"String", :width=>100, :value=>"test multi-line field data here"}]
    end

    def test_add_sheet
      @doc = Document.new self
      assert_equal [], @doc.sheets
      @doc.add_sheet "new sheet", @headers
      assert_equal 1, @doc.sheets.size
      assert_equal 'new sheet', @doc.sheets.first.name
    end
    
    def test_add_sheet_with_forward_slash_in_name
      @doc = Document.new self
      assert_equal [], @doc.sheets
      @doc.add_sheet "new sheet with /", @headers
      assert_equal 1, @doc.sheets.size
      assert_equal 'new sheet with', @doc.sheets.first.name
    end
  end
end
