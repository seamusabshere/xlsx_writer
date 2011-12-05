module XlsxWriter
  class Autofilter < ::Struct.new(:range)
    def to_xml
      %{<autoFilter ref="#{range}" />}
    end
  end
end
