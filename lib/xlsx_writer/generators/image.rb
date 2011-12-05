require 'fileutils'
module XlsxWriter
  class Image < ::Struct.new(:document, :original_path, :width, :height, :lcr, :croptop, :cropleft)
    
    AUTO = false
    
    def to_xml
      <<-EOS
<v:shape id="#{id}" o:spid="#{o_spid}" type="#_x0000_t75" style="position:absolute;margin-left:0;margin-top:0;width:#{width}pt;height:#{height}pt;z-index:1">
  <v:imagedata o:relid="#{rid}" o:title="#{o_title}" croptop=#{croptop} cropleft=#{cropleft}/>
  <o:lock v:ext="edit" rotation="t"/>
</v:shape>
EOS
    end
    
    def croptop
      self[:croptop] || 0
    end
    
    def cropleft
      self[:cropleft] || 0
    end
    
    def id
      if lcr
        lcr.image_id
      else
        o_spid #?
      end
    end
    
    def generate
      ::FileUtils.cp original_path, staging_path
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

    private

    def relative_path
      "xl/media/image#{ndx}.emf"
    end
    
    def staging_path
      p = ::File.join document.staging_dir, relative_path
      ::FileUtils.mkdir_p ::File.dirname(p)
      p
    end
  end
end
