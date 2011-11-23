require 'erb'
require 'fileutils'

module XlsxWriter
  class Xml
    attr_reader :document
    
    def initialize(document)
      @document = document
    end
    
    def path
      generate unless generated?
      @path
    end
    
    def generated?
      @generated == true
    end
    
    def staging_path
      p = ::File.join document.staging_dir, relative_path
      ::FileUtils.mkdir_p ::File.dirname(p)
      p
    end
    
    def template_path
      ::File.expand_path "../parts/#{self.class.name.demodulize.underscore}.erb", __FILE__
    end
    
    def render
      ::ERB.new(::File.read(template_path), nil, '<>').result(binding)
    end
    
    def generate
      @path = staging_path
      ::File.open(@path, 'wb') do |out|
        out.write render
      end
      Utils.unix2dos @path
      @generated = true
    end
  end
end
