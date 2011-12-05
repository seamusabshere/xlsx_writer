module XlsxWriter
  class HeaderFooter < ::Struct.new(:document, :header, :footer)
    def header
      self[:header] ||= H.new self
    end
    
    def footer
      self[:footer] ||= F.new self
    end
    
    def to_xml
      lines = []
      lines << %{<headerFooter>}
      lines << header.to_xml
      lines << footer.to_xml
      lines << %{</headerFooter>}
      if header.has_image? or footer.has_image?
        lines << %{<legacyDrawingHF r:id="rId1"/>}
      end
      lines.join("\n")
    end
    
    class HF < ::Struct.new(:header_footer, :left, :center, :right)
      def left
        self[:left] ||= L.new self
      end
      
      def center
        self[:center] ||= C.new self
      end
      
      def right
        self[:right] ||= R.new self
      end
      
      def hf
        self.class.name.demodulize
      end
      
      def to_xml
        %{<#{tag}>#{parts.map(&:to_s).join}</#{tag}>}
      end
      
      def parts
        [left,center,right].select(&:present?)
      end

      def has_image?
        parts.any?(&:has_image?)
      end
      
      class LCR < ::Struct.new(:hf, :contents)
        FONT = %{"Arial,Regular"}
        SIZE = 10

        def present?
          contents.present?
        end
        
        def has_image?
          ::Array.wrap(contents).any? { |v| v.is_a?(XlsxWriter::Image) }
        end
        
        def lcr
          self.class.name.demodulize
        end
        
        def image_id
          [ lcr, hf.hf ].join
        end
        
        def render
          out = case contents
          when :page_x_of_y
            'Page &amp;P of &amp;N'
          when ::Array
            contents.map do |v|
              case v
              when XlsxWriter::Image
                v.lcr = self
                '&amp;G'
              else
                v
              end
            end.join
          when XlsxWriter::Image
            contents.lcr = self
            '&amp;G'
          else
            contents
          end
          "K000000#{out}"
        end
        
        def to_s
          [ '', lcr, FONT, SIZE, render ].join('&amp;')
        end
      end
      
      class L < LCR; end
      class C < LCR; end
      class R < LCR; end
    end

    class H < HF
      def tag
        'oddHeader'
      end
    end
    class F < HF
      def tag
        'oddFooter'
      end
    end
  end
end
