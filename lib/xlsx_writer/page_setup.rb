class XlsxWriter
  class PageSetup
    DEFAULT = {
      :top            => 1.0,
      :right          => 0.75,
      :bottom         => 1.0,
      :left           => 0.75,
      :header         => 0.5,
      :footer         => 0.5,
      :orientation    => 'landscape',
      :vertical_dpi   => 4294967292,
      :horizontal_dpi => 4294967292
    }

    DEFAULT.keys.each do |attr|
      attr_writer attr
      define_method attr do
        instance_variable_get(:"@#{attr}") || DEFAULT[attr]
      end
    end
    
    def to_xml
      %{<pageMargins left="#{left}" right="#{right}" top="#{top}" bottom="#{bottom}" header="#{header}" footer="#{footer}"/><pageSetup orientation="#{orientation}" horizontalDpi="#{horizontal_dpi}" verticalDpi="#{vertical_dpi}"/>}
    end
  end
end
