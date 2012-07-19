require 'fast_xs'

class XlsxWriter
  class Cell
    class << self
      # 0 -> A (zero based!)
      def column_letter(i)
        result = []
        while i >= 26 do
          result << ABC[i % 26]
          i /= 26
        end
        result << ABC[result.empty? ? i : i - 1]
        result.reverse.join
      end

      # backwards compatibility
      alias :excel_column_letter :column_letter

      def type(value, proposed = nil)
        hint = if proposed
          proposed
        elsif value.is_a?(String) and value =~ TRUE_FALSE_PATTERN
          :Boolean
        else
          value.class.name.to_sym
        end
        case hint
        when :NilClass
          :String
        when :Fixnum
          :Integer
        when :Float, :Rational, :BigDecimal
          :Decimal
        when :TrueClass, :FalseClass
          :Boolean
        else
          hint
        end
      end

      def style_number(type, faded = false)
        style_number = STYLE_NUMBER[type] or raise("Don't know style number for #{type.inspect}. Must be #{STYLE_NUMBER.keys.map(&:inspect).join(', ')}.")
        if faded
          style_number * 2 + 1
        else
          style_number * 2
        end
      end

      def type_name(type)
        TYPE_NAME[type] or raise "Don't know type name for #{type.inspect}. Must be #{TYPE_NAME.keys.map(&:inspect).join(', ')}."
      end
      
      # width = Truncate([{Number of Characters} * {Maximum Digit Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256
      # Using the Calibri font as an example, the maximum digit width of 11 point font size is 7 pixels (at 96 dpi). In fact, each digit is the same width for this font. Therefore if the cell width is 8 characters wide, the value of this attribute shall be Truncate([8*7+5]/7*256)/256 = 8.7109375.
      def pixel_width(value, type = nil)
        if (w = ((character_width(value, type).to_f*MAX_DIGIT_WIDTH+5)/MAX_DIGIT_WIDTH*256)/256) < MAX_REASONABLE_WIDTH
          w
        else
          MAX_REASONABLE_WIDTH
        end
      end

      def character_width(value, type = nil)
        if type.nil?
          type = Cell.type(value)
        end
        case type
        when :String, :Integer
          value.to_s.length
        when :Decimal
          # -1000000.5
          round(value, 2).to_s.length + 2
        when :Currency
          # (1,000,000.50)
          len = round(value, 2).to_s.length + log_base(value.abs, 1e3).floor
          len += 2 if value < 0
          len
        when :Date
          DATE_LENGTH
        when :Boolean
          BOOLEAN_LENGTH
        else
          raise "Don't know character width for #{type.inspect}."
        end
      end

      def escape(value, type = nil)
        if type.nil?
          type = Cell.type(value)
        end
        case type
        when :Integer
          value.to_s
        when :Decimal, :Currency
          case value
          when BIG_DECIMAL
            value.to_s('F')
          when Rational
            value.to_f.to_s
          else
            value.to_s
          end
        when :Date
          # doesn't work for DateTimes or Times yet
          if value.is_a?(String)
            ((Time.parse(str) - JAN_1_1900) / 86_400).round
          elsif value.respond_to?(:to_date)
            (value.to_date - JAN_1_1900.to_date).to_i
          end
        when :Boolean
          value ? 1 : 0
        else
          value.fast_xs
        end
      end

      if RUBY_VERSION >= '1.9'
        def round(number, precision)
          number.round precision
        end
        def log_base(number, base)
          Math.log number, base
        end
      else
        def round(number, precision)
          (number * (10 ** precision).to_i).round / (10 ** precision).to_f
        end
        # http://blog.vagmim.com/2010/01/logarithm-to-any-base-in-ruby.html
        def log_base(number, base)
          Math.log(number) / Math.log(base)
        end
      end
    end

    ABC = ('A'..'Z').to_a
    MAX_DIGIT_WIDTH = 5
    MAX_REASONABLE_WIDTH = 75
    DATE_LENGTH = 'YYYY-MM-DD'.length
    BOOLEAN_LENGTH = 'FALSE'.length + 1
    JAN_1_1900 = Time.parse('1899-12-30 00:00:00 UTC')
    TRUE_FALSE_PATTERN = %r{^true|false$}i
    BIG_DECIMAL = defined?(BigDecimal) ? BigDecimal : Struct.new

    STYLE_NUMBER = {
      :String     => 0,
      :Boolean    => 0,
      :Currency   => 1,
      :Date       => 2,
      :Integer    => 3,
      :Decimal    => 4,
    }

    TYPE_NAME = {
      :String     => :s,
      :Boolean    => :b,
      :Currency   => :n,
      :Date       => :n,
      :Integer    => :n,
      :Decimal    => :n,
    }

    attr_reader :row
    attr_reader :x
    attr_reader :y
    attr_reader :value
    attr_reader :type

    def initialize(row, raw_value, x, y)
      @row = row
      @x = x
      @y = y
      if raw_value.is_a?(Hash)
        @value = raw_value[:value]
        @type = Cell.type @value, raw_value[:type]
        @faded_query = raw_value[:faded]
      else
        @value = raw_value
        @type = Cell.type value
      end
    end

    def faded?
      @faded_query == true
    end

    def empty?
      return @empty_query if defined?(@empty_query)
      @empty_query = (value.nil? or (value.is_a?(String) and value.empty?) or (value == false and row.sheet.document.quiet_booleans?))
    end

    def to_xml
      if empty?
        %{<c r="#{Cell.column_letter(x)}#{y}" s="0" t="s" />}
      else
        %{<c r="#{Cell.column_letter(x)}#{y}" s="#{Cell.style_number(type, faded?)}" t="#{Cell.type_name(type)}"><v>#{escaped_value}</v></c>}
      end
    end

    def pixel_width
      @pixel_width ||= Cell.pixel_width value, type
    end

    def escaped_value
      @escaped_value ||= begin
        if type == :String
          row.sheet.document.shared_strings.ndx value
        else
          Cell.escape value
        end
      end
    end
  end
end
