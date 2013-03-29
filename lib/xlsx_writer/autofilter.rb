class XlsxWriter
  class Autofilter < ::Struct.new(:sheet, :range)
    def to_xml
      %{<autoFilter ref="#{range}" />}
    end

    # Sheet1!$A$1:$B$1
    def defined_name
      "#{sheet.name}!#{dollar_range}"
    end

    def dollar_range
      a = /([A-Z]+)(\d+):([A-Z]+)(\d+)/.match(range).captures.map { |c| c.prepend '$' }
      [ a.first(2).join, a.last(2).join ].join(':')
    end
  end
end
