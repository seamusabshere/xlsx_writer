# Naview
# - Adapted to use Tempfile and ZipOutputStream rather than ZipFile
# - Replaced @zip.get_output_stream with @zip.put_next_entry
# - Commented out @zip.mkdir
# 
# source: http://github.com/harvesthq/simple_xlsx_writer
# see also: http://info.michael-simons.eu/2008/01/21/using-rubyzip-to-create-zip-files-on-the-fly/

module SimpleXlsx
  class Serializer
    def initialize tempfile
      Zip::ZipOutputStream.open(tempfile.path) do |zip|
        @zip = zip
        add_doc_props
        add_relationship_part
        add_styles
        @doc = Document.new(self)
        yield @doc
        add_workbook_relationship_part
        add_content_types
        add_workbook_part
      end
    end

    def add_workbook_part
      @zip.put_next_entry("xl/workbook.xml")
      f = @zip
      f.puts <<-ends
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<workbookPr date1904="0" />
<sheets>
ends
      @doc.sheets.each_with_index do |sheet, ndx|
        f.puts "<sheet name=\"#{sheet.name}\" sheetId=\"#{ndx + 1}\" r:id=\"#{sheet.rid}\"/>"
      end
      f.puts "</sheets></workbook>"
    end

    def open_stream_for_sheet ndx
      @zip.put_next_entry("xl/worksheets/sheet#{ndx + 1}.xml")
      @zip
    end

    def add_content_types
      @zip.put_next_entry("[Content_Types].xml")
      f = @zip
      f.puts '<?xml version="1.0" encoding="UTF-8"?>'
      f.puts '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      f.puts <<-ends
  <Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
ends
      if @doc.has_shared_strings?
        f.puts '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
      end
      @doc.sheets.each_with_index do |sheet, ndx|
        f.puts "<Override PartName=\"/xl/worksheets/sheet#{ndx+1}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
      end
      f.puts '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
      f.puts "</Types>"
    end

    def add_workbook_relationship_part
      @zip.put_next_entry("xl/_rels/workbook.xml.rels")
      f = @zip
      f.puts <<-ends
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
ends
      cnt = 0
      f.puts "<Relationship Id=\"rId#{cnt += 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>"
      @doc.sheets.each_with_index do |sheet, ndx|
        sheet.rid = "rId#{cnt += 1}"
        f.puts "<Relationship Id=\"#{sheet.rid}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet#{ndx + 1}.xml\"/>"
      end
      if @doc.has_shared_strings?
        f.puts '<Relationship Id="rId#{cnt += 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="xl/sharedStrings.xml"/>'
      end
      f.puts "</Relationships>"
    end

    def add_relationship_part
      @zip.put_next_entry("_rels/.rels")
      f = @zip
      f.puts <<-ends
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
ends
      f.puts "</Relationships>"
    end

    def add_doc_props
      @zip.put_next_entry("docProps/core.xml")
      f = @zip
      f.puts <<-ends
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
   <dcterms:created xsi:type="dcterms:W3CDTF">#{Time.now.utc.xmlschema}</dcterms:created>
   <cp:revision>0</cp:revision>
</cp:coreProperties>
ends
      @zip.put_next_entry("docProps/app.xml")
      f = @zip
      f.puts <<-ends
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <TotalTime>0</TotalTime>
</Properties>
ends
    end

    def add_styles
      @zip.put_next_entry("xl/styles.xml")
      f = @zip
      f.puts <<-ends
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0.00"/>
  </numFmts>
  <fonts count="2">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="minor"/>
    </font>
    <font>
      <b/>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="minor"/>
    </font>
  </fonts>
  <fills count="3">
    <fill>
      <patternFill patternType="none"/>
    </fill>
    <fill>
      <patternFill patternType="gray125"/>
    </fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor theme="0" tint="-0.14999847407452621"/>
        <bgColor indexed="64"/>
      </patternFill>
    </fill>
  </fills>
  <borders count="1">
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="7">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="1" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="0" fontId="1" fillId="1" borderId="0" xfId="0" applyFont="1" applyFill="1" applyProtection="1"/>
    <xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1" applyProtection="1"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
ends
    end
  end
end

