require 'fileutils'

module XlsxWriter
  class Document
    class << self
      def auto
        ::Dir[::File.expand_path('../generators/*.rb', __FILE__)].map do |path|
          XlsxWriter.const_get ::File.basename(path, '.rb').camelcase
        end.reject do |klass|
          klass.const_defined?(:AUTO) and klass.const_get(:AUTO) == false
        end
      end
    end
    
    def add_sheet(name)
      raise ::RuntimeError, "Can't add sheet, already generated!" if generated?
      sheet = Sheet.new self, name
      sheets << sheet
      sheet
    end
    
    def page_setup
      @page_setup ||= PageSetup.new
    end
    
    def header_footer
      @header_footer ||= HeaderFooter.new self
    end
    
    delegate :header, :footer, :to => :header_footer
    
    def add_image(path, width, height)
      raise ::RuntimeError, "Can't add image, already generated!" if generated?
      image = Image.new self, path, width, height
      images << image
      image
    end

    def path
      generate unless generated?
      @path
    end

    def cleanup
      ::File.unlink(@path) if ::File.exist?(@path)
      ::FileUtils.rm_rf(@staging_dir) if ::File.exist?(@staging_dir)
      @path = nil
      @staging_dir = nil
      @generated = false
    end

    def sheets #:nodoc:
      @sheets ||= []
    end
    
    def images
      @images ||= []
    end
    
    def staging_dir
      @staging_dir ||= Utils.tmp_path
      ::FileUtils.mkdir_p @staging_dir
      @staging_dir
    end

    private
    
    def generate
      sheets.each(&:generate)
      images.each(&:generate)
      Document.auto.each do |part|
        part.new(self).generate
      end
      @path = Utils.zip staging_dir
      @generated = true
    end

    def generated?
      @generated == true
    end
  end
end
