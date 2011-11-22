require 'fileutils'

module XlsxWriter
  class Document
    class << self
      def parts
        ::Dir[::File.expand_path('../parts/*.rb', __FILE__)].map do |path|
          XlsxWriter.const_get ::File.basename(path, '.rb').camelcase
        end
      end
    end
    
    def add_sheet(name)
      raise ::RuntimeError, "Can't add sheet, already generated!" if generated?
      sheet = Sheet.new self, name
      sheets << sheet
      sheet
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
    
    def staging_dir
      @staging_dir ||= Utils.tmp_path
      ::FileUtils.mkdir_p @staging_dir
      @staging_dir
    end

    private
    
    def generate
      sheets.each(&:generate)
      (Document.parts - [Sheet]).each do |part|
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
