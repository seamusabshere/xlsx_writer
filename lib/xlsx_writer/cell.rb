require 'fast_xs'

class XlsxWriter
  class Cell
    class << self
      # TODO make a class for this
      def excel_type(calculated_type)
        case calculated_type
        when :String
          :string
        when :Number, :Integer, :Decimal, :Date, :Currency
          :n
        when :Boolean
          :b
        else
          raise ::ArgumentError, "Unknown cell type #{calculated_type}"
        end
      end
      
      # TODO make a class for this
      def excel_style_number(calculated_type, faded = false)
        i = case calculated_type
        when :String
          0
        when :Boolean
          0 # todo
        when :Currency
          1
        when :Date
          2
        when :Number, :Integer
          3
        when :Decimal
          4
        else
          raise ::ArgumentError, "Unknown cell type #{k}"
        end
        if faded
          i * 2 + 1
        else
          i * 2
        end
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
        unless str =~ /\A[0-9\.\-]*\z/
          raise ::ArgumentError, %{Bad value "#{value}" Only numbers and dots (.) allowed in number fields}
        end
        str.fast_xs
      end
      
      alias :excel_currency :excel_number
      alias :excel_integer :excel_number
      alias :excel_decimal :excel_number
      
      # doesn't necessarily work for times yet
      def excel_date(value)
        if value.is_a?(::String)
          ((::Time.parse(str) - JAN_1_1900) / 86_400).round
        elsif value.respond_to?(:to_date)
          (value.to_date - JAN_1_1900.to_date).to_i
        end
      end
      
      def excel_boolean(value)
        value ? 1 : 0
      end

      # width = Truncate([{Number of Characters} * {Maximum Digit Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256
      # Using the Calibri font as an example, the maximum digit width of 11 point font size is 7 pixels (at 96 dpi). In fact, each digit is the same width for this font. Therefore if the cell width is 8 characters wide, the value of this attribute shall be Truncate([8*7+5]/7*256)/256 = 8.7109375.
      def pixel_width(character_width)
        [
          ((character_width.to_f*MAX_DIGIT_WIDTH+5)/MAX_DIGIT_WIDTH*256)/256,
          MAX_REASONABLE_WIDTH
        ].min
      end

      def calculate_type(value)
        case value
        when Date
          :Date
        when Integer
          :Integer
        when Float
          :Decimal
        when Numeric
          :Number
        when TrueClass, FalseClass, TRUE_FALSE_PATTERN
          :Boolean
        else
          if (defined?(Decimal) and value.is_a?(Decimal)) or (defined?(BigDecimal) and value.is_a?(BigDecimal))
            :Decimal
          else
            :String
          end
        end
      end

      def character_width(value, calculated_type = nil)
        calculated_type ||= calculate_type(value)
        case calculated_type
        when :String
          value.to_s.length
        when :Number, :Integer, :Decimal
          # -1000000.5
          len = round(value, 2).to_s.length
          len += 2 if calculated_type == :Decimal
          len += 1 if value < 0
          len
        when :Currency
          # (1,000,000.50)
          len = round(value, 2).to_s.length + log_base(value.abs, 1e3).floor
          len += 2 if value < 0
          len
        when :Date
          DATE_LENGTH
        when :Boolean
          BOOLEAN_LENGTH
        end
      end

      if ::RUBY_VERSION >= '1.9'
        def round(number, precision)
          number.round precision
        end
        def log_base(number, base)
          ::Math.log number, base
        end
      else
        def round(number, precision)
          (number * (10 ** precision).to_i).round / (10 ** precision).to_f
        end
        # http://blog.vagmim.com/2010/01/logarithm-to-any-base-in-ruby.html
        def log_base(number, base)
          ::Math.log(number) / ::Math.log(base)
        end
      end
    end
    
    ABC = ('A'..'Z').to_a
    MAX_DIGIT_WIDTH = 5
    MAX_REASONABLE_WIDTH = 75
    DATE_LENGTH = 'YYYY-MM-DD'.length
    BOOLEAN_LENGTH = 'FALSE'.length + 1
    JAN_1_1900 = ::Time.parse('1899-12-30 00:00:00 UTC')
    TRUE_FALSE_PATTERN = %r{^true|false$}i
    
    attr_reader :row
    attr_reader :value
    attr_reader :pixel_width
    attr_reader :excel_type
    attr_reader :excel_style_number
    attr_reader :excel_value

    def initialize(row, data)
      @row = row
      if data.is_a?(::Hash)
        data = data.symbolize_keys
        @value = data[:value]
        faded = data[:faded]
        calculated_type = data[:type] || Cell.calculate_type(@value)
      else
        @value = data
        faded = false
        calculated_type = Cell.calculate_type @value
      end
      character_width = Cell.character_width @value, calculated_type
      @pixel_width = Cell.pixel_width character_width
      @excel_type = Cell.excel_type calculated_type
      @excel_style_number = Cell.excel_style_number calculated_type, faded
      @excel_value = Cell.send "excel_#{calculated_type.to_s.underscore}", @value
    end

    def to_xml
      if value.nil? or (value.is_a?(String) and value.empty?) or (value == false and quiet_booleans?)
        %{<c r="#{excel_column_letter}#{row.ndx}" s="0" t="s" />}
      elsif excel_type == :string
        unless document.shared_strings.has_key?(value)
          document.shared_strings[value] = document.shared_strings.count
        end

        string_index = document.shared_strings[value]
        %{<c r="#{excel_column_letter}#{row.ndx}" s="#{excel_style_number}" t="s"><v>#{string_index}</v></c>}
      else
        %{<c r="#{excel_column_letter}#{row.ndx}" s="#{excel_style_number}" t="#{excel_type}"><v>#{excel_value}</v></c>}
      end
    end

    # 0 -> A (zero based!)
    def excel_column_letter
      Cell.excel_column_letter row.cells.index(self)
    end

    private

    def document
      row.sheet.document
    end

    def quiet_booleans?
      return @quiet_booleans if defined?(@quiet_booleans)
      @quiet_booleans = row.sheet.document.quiet_booleans?
    end
  end
end
