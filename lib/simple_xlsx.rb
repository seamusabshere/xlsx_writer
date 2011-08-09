require 'tempfile'
require 'rubygems'

$:.unshift(File.dirname(__FILE__))
require 'simple_xlsx/xml_escape'
require 'simple_xlsx/monkey_patches_for_true_zip_stream'
require 'simple_xlsx/serializer'
require 'simple_xlsx/document'
require 'simple_xlsx/sheet'


# add lf -> crlf conversion to string
class String
  def lf_to_crlf
    gsub(/\012/, "\015\012")
  end
end

