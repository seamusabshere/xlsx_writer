# Naview
# - Adapted to use Tempfile and ZipOutputStream rather than ZipFile
# - Replaced @zip.get_output_stream with @zip.put_next_entry
# - Commented out @zip.mkdir
# 
# source: http://github.com/harvesthq/simple_xlsx_writer
# see also: http://info.michael-simons.eu/2008/01/21/using-rubyzip-to-create-zip-files-on-the-fly/

module SimpleXlsx
  class Serializer
    def initialize filename
      @base_dir = Dir.mktmpdir
      add_doc_props
      add_relationship_part
      add_styles
      @doc = Document.new(self)
      yield @doc
      add_workbook_relationship_part
      add_content_types
      add_workbook_part
      zip_command = "zip -r -q #{filename} ."
      cwd = Dir.getwd
      Dir.chdir(@base_dir)
      result = system zip_command
      if result
        @zipfile = "#{@base_dir}/#{filename}"
      else
        @zipfile = nil
      end
      Dir.chdir(cwd)
    end

    def cleanup
      if File.exists?(@base_dir)
        FileUtils.rm_r(@base_dir)
      end
    end
    
    def zipfile
      @zipfile
    end
    
    def add_workbook_part
      workbook_dir = "#{@base_dir}/xl/"
      FileUtils.mkdir_p(workbook_dir) unless File.exists?(workbook_dir)
      filename = "#{workbook_dir}/workbook.xml"
      f = File.open(filename, "wb")
      f.puts <<-ends.lf_to_crlf
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <workbookPr date1904="0" />
    <sheets>
ends
      @doc.sheets.each_with_index do |sheet, ndx|
        f.puts <<-ends.lf_to_crlf
      <sheet name="#{sheet.name}" sheetId="#{ndx + 1}" r:id="#{sheet.rid}"/>
ends
      end
      f.puts <<-ends.lf_to_crlf
    </sheets>
  </workbook>
ends
      f.close
    end

    def open_stream_for_sheet ndx
      sheet_dir = "#{@base_dir}/xl/worksheets/"
      FileUtils.mkdir_p(sheet_dir) unless File.exists?(sheet_dir)
      filename = "#{sheet_dir}/sheet#{ndx + 1}.xml"
      File.open(filename, "wb")
    end

    def add_content_types
      filename = "#{@base_dir}/[Content_Types].xml"
      f = File.open(filename, "wb")
      f.puts <<-ends.lf_to_crlf
<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
ends
      if @doc.has_shared_strings?
        f.puts <<-ends.lf_to_crlf
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
ends
      end
      @doc.sheets.each_with_index do |sheet, ndx|
        f.puts <<-ends.lf_to_crlf
  <Override PartName="/xl/worksheets/sheet#{ndx+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
ends
      end
      f.puts <<-ends.lf_to_crlf
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
ends
      f.close
    end

    def add_workbook_relationship_part
      workbook_rel_dir = "#{@base_dir}/xl/_rels"
      FileUtils.mkdir_p(workbook_rel_dir) unless File.exists?(workbook_rel_dir)
      filename = "#{workbook_rel_dir}/workbook.xml.rels"
      f = File.open(filename, "wb")
      f.puts <<-ends.lf_to_crlf
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
ends
      cnt = 0
      f.puts <<-ends.lf_to_crlf
  <Relationship Id="rId#{cnt += 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
ends
      @doc.sheets.each_with_index do |sheet, ndx|
        sheet.rid = "rId#{cnt += 1}"
        f.puts <<-ends.lf_to_crlf
  <Relationship Id="#{sheet.rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet#{ndx + 1}.xml"/>
ends
      end
      if @doc.has_shared_strings?
        f.puts <<-ends.lf_to_crlf
  <Relationship Id="rId#{cnt += 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="xl/sharedStrings.xml"/>
ends
      end
      f.puts <<-ends.lf_to_crlf
</Relationships>
ends
      f.close
    end

    def add_relationship_part
      reldir = "#{@base_dir}/_rels"
      FileUtils.mkdir_p(reldir) unless File.exists?(reldir)
      filename = "#{reldir}/.rels"
      f = File.open(filename, "wb")
      f.puts <<-ends.lf_to_crlf
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
ends
      f.close
    end

    def add_doc_props
      propsdir = "#{@base_dir}/docProps"
      FileUtils.mkdir_p(propsdir) unless File.exists?(propsdir)
      filename = "#{propsdir}/core.xml"
      f = File.open(filename, "wb")
      f.puts <<-ends.lf_to_crlf
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
   <dcterms:created xsi:type="dcterms:W3CDTF">#{Time.now.utc.xmlschema}</dcterms:created>
   <cp:revision>0</cp:revision>
</cp:coreProperties>
ends
      f.close
      filename = "#{propsdir}/app.xml"
      f = File.open(filename, "wb")
      f.puts <<-ends.lf_to_crlf
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <TotalTime>0</TotalTime>
</Properties>
ends
      f.close
    end

    def add_styles
      styledir = "#{@base_dir}/xl"
      FileUtils.mkdir_p(styledir) unless File.exists?(styledir)
      filename = "#{styledir}/styles.xml"
      f = File.open(filename, "wb")
      f.puts <<-ends.lf_to_crlf
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
    <xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyProtection="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
ends
      f.close
    end
  end
end

