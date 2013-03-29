require 'fileutils'

class XlsxWriter
  class Image
    DEFAULT = {
      :croptop => 0,
      :cropleft => 0
    }
    AUTO = false

    attr_reader :document
    attr_reader :original_path
    attr_reader :width
    attr_reader :height
    attr_accessor :lcr
    attr_writer :croptop
    attr_writer :cropleft

    def initialize(document, original_path, width, height)
      @document = document
      @original_path = original_path
      @width = width
      @height = height
      @mutex = ::Mutex.new
    end

    def to_xml
      <<-EOS
<v:shape id="#{id}" o:spid="#{o_spid}" type="#_x0000_t75" style="position:absolute;margin-left:0;margin-top:0;width:#{width}pt;height:#{height}pt;z-index:1">
  <v:imagedata o:relid="#{rid}" o:title="#{o_title}" croptop="#{croptop}" cropleft="#{cropleft}"/>
  <o:lock v:ext="edit" rotation="t"/>
</v:shape>
EOS
    end
    
    def croptop
      @croptop || DEFAULT[:croptop]
    end
    
    def cropleft
      @cropleft || DEFAULT[:cropleft]
    end
    
    def id
      if lcr
        lcr.image_id
      else
        o_spid #?
      end
    end

    def generate
      path
      true
    end

    def generated?
      @generated == true
    end

    def path
      @path || @mutex.synchronize do
        @path ||= begin
          memo = ::File.join document.staging_dir, relative_path
          ::FileUtils.mkdir_p ::File.dirname(memo)
          ::FileUtils.cp original_path, memo
          @generated = true
          memo
        end
      end
    end
    
    def ndx
      document.images.index(self) + 1
    end
    
    def rid
      "rId#{ndx}"
    end
    
    def o_title
      ::File.basename(original_path)
    end
    
    def o_spid
      "_x0000_s#{1025+ndx}"
    end
    
    def absolute_path
      "/#{relative_path}"
    end

    def relative_path
      "xl/media/image#{ndx}.emf"
    end
  end
end
