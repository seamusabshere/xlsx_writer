# encoding: UTF-8
require 'helper'

describe XlsxWriter do
  describe "hello world example" do
    before do
      @doc = XlsxWriter::Document.new
      @sheet1 = @doc.add_sheet("People")
      @sheet1.add_row(['header1', 'header2'])
      @sheet1.add_row(['hello', 'world'])
      @sheet1.add_row(['affirmative', true])
      @sheet1.add_row(['negative', false])
    end
    after do
      @doc.try :cleanup
    end
    it "returns a path to an xlsx" do
      File.exist?(@doc.path).must_equal true
      File.extname(@doc.path).must_equal '.xlsx'
    end
    it "is a readable xlsx" do
      RemoteTable.new(@doc.path, :format => :xlsx).rows.first.must_equal('header1' => 'hello', 'header2' => 'world')
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

  describe "quiet booleans" do
    it "shows TRUE or FALSE for booleans by default" do
      doc = XlsxWriter::Document.new
      sheet1 = doc.add_sheet("QuietBooleans")
      sheet1.add_row(['affirmative', 'negative'])
      sheet1.add_row([true, false])
      t = RemoteTable.new(doc.path, :format => :xlsx)
      t[0]['affirmative'].must_equal 'TRUE'
      t[0]['negative'].must_equal 'FALSE'
      doc.cleanup
    end
    it "shows TRUE or blank for false if quiet booleans is enabled" do
      doc = XlsxWriter::Document.new
      doc.quiet_booleans!
      sheet1 = doc.add_sheet("QuietBooleans")
      sheet1.add_row(['affirmative', 'negative'])
      sheet1.add_row([true, false])
      t = RemoteTable.new(doc.path, :format => :xlsx)
      t[0]['affirmative'].must_equal 'TRUE'
      t[0]['negative'].must_equal ''
      doc.cleanup
    end
  end

  describe "freeze pane under a certain row" do
    before do
      @doc = XlsxWriter::Document.new
      @sheet1 = @doc.add_sheet("Freeze")
      @sheet1.add_row([1, 2])
      @sheet1.add_row([3, 4])
    end
    after do
      @doc.try :cleanup
    end
    it "doesn't freeze by default" do
      dir = UnixUtils.unzip @doc.path
      File.read("#{dir}/xl/worksheets/sheet1.xml").wont_include 'pane'
      FileUtils.rm_rf dir
    end
    it "shows TRUE or blank for false if quiet booleans is enabled" do
      @sheet1.freeze_top_left = 'A3'
      dir = UnixUtils.unzip @doc.path
      xml = File.read("#{dir}/xl/worksheets/sheet1.xml")
      xml.must_include 'topLeftCell="A3"'
      xml.must_include 'ySplit="2"'
      FileUtils.rm_rf dir
    end
  end

  describe "numeric types" do
    before do
      @reference = {
        # 'Integer' => [1, '1'], - remote_table always converts to 1.0
        'Float' => [1.2345, '1.2345'],
        'Decimal' => [Decimal('6.789'), '6.789'],
        'BigDecimal' => [::BigDecimal.new(9.876, 4), '9.876'],
        'Rational' => [Rational(2, 3), '0.6666'],
      }
      @doc = XlsxWriter::Document.new
      @sheet1 = @doc.add_sheet("Numbers")
      @reference.each do |k, v|
        @sheet1.add_row([k, v.first])
      end
    end
    after do
      @doc.try :cleanup
    end
    it "renders different numeric types into plain decimals" do
      t = RemoteTable.new(@doc.path, :format => :xlsx, :headers => %w{ type number })
      t.each do |row|
        row['number'].to_s.first(6).must_equal @reference[row['type']].last, "for #{row['type']}"
      end
    end
  end

  describe "example with autofilter, header image, and footer text" do
    before do
      @doc = XlsxWriter::Document.new
      sheet1 = @doc.add_sheet("People")
      # DATA
      sheet1.add_row(%w{DoB Name Occupation Number Integer Float})
      sheet1.add_row([
        Date.parse("July 31, 1912"), 
        "Milton Friedman",
        "Economist / Statistician",
        {:type => :Currency, :value => 99_000_000},
        500_000,
        500_000.00,
      ])
      sheet1.add_autofilter 'A1:F1'
      # FORMATTING
      @doc.page_setup.top = 1.5
      # hint: set up your header/footer in Excel, save, unzip the xlsx, get the .emf files, croptop, etc. from there
      left_header_image = @doc.add_image(File.expand_path('../support/image1.emf', __FILE__), 118, 107)
      left_header_image.croptop = '11025f'
      left_header_image.cropleft = '9997f'
      center_footer_image = @doc.add_image(File.expand_path('../support/image2.emf', __FILE__), 116, 36)
      @doc.page_setup.header = 0
      @doc.page_setup.footer = 0
      @doc.header.left.contents = left_header_image
      @doc.header.right.contents = 'Reporting Program'
      @doc.footer.center.contents = [ 'Powered by ', center_footer_image ]
      @doc.footer.right.contents = :page_x_of_y
      @unzipped = UnixUtils.unzip @doc.path
    end
    after do
      @doc.try :cleanup
      FileUtils.rm_rf @unzipped if @unzipped
    end
    it "has an autofilter" do
      File.read("#{@unzipped}/xl/worksheets/sheet1.xml").must_include %{<autoFilter ref="A1:F1" />}
    end
    it "has a header image" do
      File.read("#{@unzipped}/xl/drawings/_rels/vmlDrawing1.vml.rels").must_include %{<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="/xl/media/image1.emf"/>}
      File.read("#{@unzipped}/xl/drawings/vmlDrawing1.vml").must_include %{<v:imagedata o:relid="rId1" o:title="image1.emf" croptop="11025f" cropleft="9997f"/>}
      original = UnixUtils.md5sum File.expand_path("../support/image1.emf", __FILE__)
      UnixUtils.md5sum("#{@unzipped}/xl/media/image1.emf").must_equal original
    end
  end

  describe 'Cell' do
    describe :round do
      it "works same in 1.8 and 1.9" do
        XlsxWriter::Cell.round(12.3456, 1).must_equal 12.3
        XlsxWriter::Cell.round(12.3456, 2).must_equal 12.35
      end
    end
    describe :log_base do
      it "works same in 1.8 and 1.9" do
        XlsxWriter::Cell.log_base(4, 1e3).must_be_close_to 0.2
        XlsxWriter::Cell.log_base(4, 12.345).must_be_close_to 0.552
      end
    end
  end
end
