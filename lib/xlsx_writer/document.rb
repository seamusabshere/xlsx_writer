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

    attr_reader :staging_dir
    attr_reader :sheets
    attr_reader :images
    attr_reader :page_setup
    attr_reader :header_footer
    attr_reader :shared_strings

    def initialize
      staging_dir = ::UnixUtils.tmp_path 'xlsx_writer'
      ::FileUtils.mkdir_p staging_dir
      @staging_dir = staging_dir
      @sheets = []
      @images = []
      @page_setup = PageSetup.new
      @header_footer = HeaderFooter.new
      @shared_strings = {}
      @mutex = ::Mutex.new
    end

    # Instead of TRUE or FALSE, show TRUE and blank if false
    def quiet_booleans!
      @quiet_booleans = true
    end

    def quiet_booleans?
      @quiet_booleans == true
    end

    def add_sheet(name)
      raise ::RuntimeError, "Can't add sheet, already generated!" if generated?
      sheet = Sheet.new self, name
      sheets << sheet
      sheet
    end
    
    delegate :header, :footer, :to => :header_footer
    
    def add_image(path, width, height)
      raise ::RuntimeError, "Can't add image, already generated!" if generated?
      image = Image.new self, path, width, height
      images << image
      image
    end

    def path
      @path || @mutex.synchronize do
        @path ||= begin 
          sheets.each(&:generate)
          images.each(&:generate)
          Document.auto.each do |part|
            part.new(self).generate
          end
          with_zip_extname = ::UnixUtils.zip staging_dir
          with_xlsx_extname = with_zip_extname.sub(/.zip$/, '.xlsx')
          ::FileUtils.mv with_zip_extname, with_xlsx_extname
          @generated = true
          with_xlsx_extname
        end
      end
    end

    def cleanup
      ::FileUtils.rm_rf @staging_dir.to_s
      ::FileUtils.rm_f @path.to_s
    end
    
    def generate
      path
      true
    end
    
    def generated?
      @generated == true
    end
  end
end
