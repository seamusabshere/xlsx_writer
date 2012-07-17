class XlsxWriter
  class HeaderFooter
    attr_reader :header
    attr_reader :footer

    def initialize
      @header = HF.new 'H', 'oddHeader'
      @footer = HF.new 'F', 'oddFooter'
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
    
    class HF
      attr_reader :id
      attr_reader :tag
      attr_reader :left
      attr_reader :center
      attr_reader :right
      
      def initialize(id, tag)
        @id = id
        @tag = tag
        @left = LCR.new self, 'L'
        @center = LCR.new self, 'C'
        @right = LCR.new self, 'R'
      end
      
      def to_xml
        %{<#{tag}>#{parts.map(&:code).join}</#{tag}>}
      end
      
      def parts
        [left,center,right].select(&:present?)
      end

      def has_image?
        parts.any?(&:has_image?)
      end
      
      class LCR
        FONT = %{"Arial,Regular"}
        SIZE = 10

        attr_accessor :contents
        attr_reader :hf
        attr_reader :id

        def initialize(hf, id)
          @hf = hf
          @id = id
        end

        def present?
          contents.present?
        end
        
        def has_image?
          ::Array.wrap(contents).any? { |v| v.is_a?(XlsxWriter::Image) }
        end
        
        def image_id
          [ id, hf.id ].join
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
        
        def code
          [ '', id, FONT, SIZE, render ].join('&amp;')
        end
      end
    end
  end
end
