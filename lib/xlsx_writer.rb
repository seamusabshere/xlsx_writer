require 'thread'
require 'active_support/core_ext'
require 'unix_utils'

module XlsxWriter
end

require 'xlsx_writer/cell'
require 'xlsx_writer/document'
require 'xlsx_writer/row'
require 'xlsx_writer/header_footer'
require 'xlsx_writer/autofilter'
require 'xlsx_writer/page_setup'

require 'xlsx_writer/xml'
# manual
require 'xlsx_writer/xml/sheet'
require 'xlsx_writer/xml/sheet_rels'
require 'xlsx_writer/xml/image'
require 'xlsx_writer/xml/shared_strings'

# automatic
require 'xlsx_writer/xml/app'
require 'xlsx_writer/xml/content_types'
require 'xlsx_writer/xml/doc_props'
require 'xlsx_writer/xml/rels'
require 'xlsx_writer/xml/styles'
require 'xlsx_writer/xml/workbook'
require 'xlsx_writer/xml/workbook_rels'
require 'xlsx_writer/xml/vml_drawing'
require 'xlsx_writer/xml/vml_drawing_rels'
