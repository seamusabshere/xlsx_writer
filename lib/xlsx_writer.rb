require 'active_support/core_ext'

require 'xlsx_writer/version'

module XlsxWriter
  def self.gem_dir
    ::File.join ::File.dirname(__FILE__), 'xlsx_writer'
  end
  
  autoload :Cell, "#{gem_dir}/cell"
  autoload :Document, "#{gem_dir}/document"
  autoload :Row, "#{gem_dir}/row"
  autoload :Utils, "#{gem_dir}/utils"
  autoload :Xml, "#{gem_dir}/xml"
  autoload :HeaderFooter, "#{gem_dir}/header_footer"
  autoload :Autofilter, "#{gem_dir}/autofilter"
  autoload :PageSetup, "#{gem_dir}/page_setup"
  
  # manual
  autoload :Sheet, "#{gem_dir}/generators/sheet"
  autoload :SheetRels, "#{gem_dir}/generators/sheet_rels"
  autoload :Image, "#{gem_dir}/generators/image"
  
  # generators
  autoload :App, "#{gem_dir}/generators/app"
  autoload :ContentTypes, "#{gem_dir}/generators/content_types"
  autoload :DocProps, "#{gem_dir}/generators/doc_props"
  autoload :Rels, "#{gem_dir}/generators/rels"
  autoload :Styles, "#{gem_dir}/generators/styles"
  autoload :Workbook, "#{gem_dir}/generators/workbook"
  autoload :WorkbookRels, "#{gem_dir}/generators/workbook_rels"
  autoload :VmlDrawing, "#{gem_dir}/generators/vml_drawing"
  autoload :VmlDrawingRels, "#{gem_dir}/generators/vml_drawing_rels"
end
