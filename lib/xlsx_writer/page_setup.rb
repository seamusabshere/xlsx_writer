module XlsxWriter
  class PageSetup < ::Struct.new(:top, :right, :bottom, :left, :header, :footer, :orientation, :vertical_dpi, :horizontal_dpi)
    def top
      self[:top] || 1.0
    end
    
    def right
      self[:right] || 0.75
    end
    
    def bottom
      self[:bottom] || 1.0
    end
    
    def left
      self[:left] || 0.75
    end
    
    def header
      self[:header] || 0.5
    end
    
    def footer
      self[:footer] || 0.5
    end
    
    def orientation
      self[:orientation] || 'landscape'
    end
    
    def vertical_dpi
      self[:vertical_dpi] || 4294967292
    end
    
    def horizontal_dpi
      self[:horizontal_dpi] || 4294967292
    end
    
    def to_xml
      %{<pageMargins left="#{left}" right="#{right}" top="#{top}" bottom="#{bottom}" header="#{header}" footer="#{footer}"/><pageSetup orientation="#{orientation}" horizontalDpi="#{horizontal_dpi}" verticalDpi="#{vertical_dpi}"/>}
    end
  end
end
