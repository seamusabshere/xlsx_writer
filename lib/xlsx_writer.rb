require 'active_support/core_ext'

module XlsxWriter
  def self.gem_dir
    ::File.join ::File.dirname(__FILE__), 'xlsx_writer'
  end
  
  autoload :Cell, "#{gem_dir}/cell"
  autoload :Document, "#{gem_dir}/document"
  autoload :Row, "#{gem_dir}/row"
  autoload :Utils, "#{gem_dir}/utils"
  autoload :Xml, "#{gem_dir}/xml"
  
  # parts
  autoload :App, "#{gem_dir}/parts/app"
  autoload :ContentTypes, "#{gem_dir}/parts/content_types"
  autoload :DocProps, "#{gem_dir}/parts/doc_props"
  autoload :Rels, "#{gem_dir}/parts/rels"
  autoload :Sheet, "#{gem_dir}/parts/sheet"
  autoload :Styles, "#{gem_dir}/parts/styles"
  autoload :Workbook, "#{gem_dir}/parts/workbook"
  autoload :WorkbookRels, "#{gem_dir}/parts/workbook_rels"
end
