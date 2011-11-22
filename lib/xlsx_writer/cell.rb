require 'fast_xs'

module XlsxWriter
  class Cell
    class << self
      def excel_type(value, type_hint = nil)
        if value.is_a?(::Date)
          return :n
        end
        
        # unless type_hint
        #   return :inlineStr
        # end
        
        case type_hint.to_sym
        when :String
          :inlineStr
        when :Number, :DateTime, :Money
          :n
        when :Boolean
          :b
        else
          raise ::ArgumentError, "Unknown cell type #{k}"
        end
      end
      
      def excel_style_number(value, type_hint = nil)
        if value.is_a?(::Date)
          return 1
        end
        
        # unless type_hint
        #   return 3
        # end

        case type_hint.to_sym
        when :String
          3
        when :Number
          6
        when :DateTime
          1
        when :Boolean
          5
        when :Money
          2
        else
          raise ::ArgumentError, "Unknown cell type #{k}"
        end if type_hint
      end
      
      def excel_column_letter(i)
        result = []
        while i >= 26 do
          result << ABC[i % 26]
          i /= 26
        end
        result << ABC[result.empty? ? i : i - 1]
        result.reverse.join
      end
            
      def excel_string(value)
        value.to_s.fast_xs
      end
      
      def excel_number(value)
        str = value.to_s.dup
        str.gsub! '$', '' # ?
        str.gsub! ',', '' # ?
        str.fast_xs
      end
      
      alias :excel_money :excel_number
      
      # http://support.microsoft.com/kb/180162
      JAN_1_1900 = ::Time.parse('1900-01-01')
      def excel_date_time(value)
        case value
        when ::String
          ((::Time.parse(str) - JAN_1_1900) / 86_400).round
        when ::Date
          (value - JAN_1_1900.to_date).to_i
        end
      end
      
      def excel_boolean(value)
        value ? 1 : 0
      end
    end
    
    ABC = ('A'..'Z').to_a
    
    attr_reader :row
    attr_reader :data
    
    def initialize(row, data)
      @row = row
      @data = data.is_a?(::Hash) ? data.symbolize_keys : data
      validate
    end
    
    def unstyled?
      !styled?
    end
    
    def styled?
      data.is_a?(::Hash)
    end
    
    def to_xml
      if value.blank?
        %{<c r="#{excel_column_letter}#{row.ndx}" s="#{excel_style_number}" t="#{excel_type}" />}
      elsif excel_type == :inlineStr
        %{<c r="#{excel_column_letter}#{row.ndx}" s="#{excel_style_number}" t="#{excel_type}"><is><t>#{excel_value}</t></is></c>}
      else
        %{<c r="#{excel_column_letter}#{row.ndx}" s="#{excel_style_number}" t="#{excel_type}"><v>#{excel_value}</v></c>}
      end
    end
    
    # 0 -> A (zero based!)
    def excel_column_letter
      Cell.excel_column_letter row.cells.index(self)
    end
        
    def validate
      if styled? and type_hint.blank?
        raise ::ArgumentError, "When passing a Hash to Sheet#add_row, must specify cell type"
      end
    end
    
    # detect dates here, even if we're not styled
    def excel_type
      Cell.excel_type value, type_hint
    end
    
    def excel_style_number
      Cell.excel_style_number value, type_hint
    end

    def type_hint
      if styled?
        data[:type]
      elsif value.is_a?(::Date)
        :DateTime
      else
        :String
      end
    end
    
    def value
      styled? ? data[:value] : data
    end
    
    def excel_value
      Cell.send "excel_#{type_hint.to_s.underscore}", value
    end
  end
end
