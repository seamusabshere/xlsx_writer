# encoding: UTF-8
require 'helper'

describe XlsxWriter do
  describe "hello world example" do
    before do
      @doc = XlsxWriter::Document.new
      @sheet1 = @doc.add_sheet("People")
      @sheet1.add_row(['header1', 'header2'])
      @sheet1.add_row(['hello', 'world'])
    end
    after do
      # @doc.cleanup
    end
    it "returns a path to an xlsx" do
      File.exist?(@doc.path).must_equal true
      File.extname(@doc.path).must_equal '.xlsx'
    end
    it "is a readable xlsx" do
      RemoteTable.new("file://#{@doc.path}", :format => :xlsx).rows.first.must_equal('header1' => 'hello', 'header2' => 'world')
    end
    it "only generates once" do
      @doc.generate
      mtime = File.mtime(@doc.path)
      md5 = UnixUtils.md5sum(@doc.path)
      @doc.generate
      File.mtime(@doc.path).must_equal mtime
      UnixUtils.md5sum(@doc.path).must_equal md5
    end
    it "won't accept new sheets once it's been generated" do
      @doc.add_sheet 'okfine'
      @doc.generate
      lambda {
        @doc.add_sheet 'toolate'
      }.must_raise(RuntimeError, /already generated/)
    end
    it "won't accept new rows once it's been generated" do
      @sheet1.add_row(['phew'])
      @doc.generate
      lambda {
        @sheet1.add_row(['too', 'late', 'to', 'apologize'])
      }.must_raise(RuntimeError, /already generated/)
    end
    it "automatically generates if you call #path" do
      @doc.generated?.must_equal false
      @doc.path
      @doc.generated?.must_equal true
    end
  end
end
