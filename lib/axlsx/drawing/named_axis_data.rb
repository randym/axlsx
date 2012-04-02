# encoding: UTF-8
# TODO: review cat, val and named access data to extend this and reduce replicated code.
module Axlsx
  # The ValAxisData class manages the values for a chart value series.
  class NamedAxisData < CatAxisData

    # creates a new NamedAxisData Object
    # @param [String] name The serialized node name for the axis data object
    # @param [Array] The data to associate with the axis data object
    def initialize(name, data=[])
      super(data)
      @name = name
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<c:' << @name.to_s << '>'
      str << '<c:numRef>'
      str << '<c:f>' << Axlsx::cell_range(@list) << '</c:f>'
      str << '<c:numCache>'
      str << '<c:formatCode>General</c:formatCode>'
      str << '<c:ptCount val="' << size.to_s << '"/>'
      each_with_index do |item, index|
        v = item.is_a?(Cell) ?  item.value.to_s : item
        str << '<c:pt idx="' << index.to_s << '"><c:v>' << v << '</c:v></c:pt>'
      end
      str << '</c:numCache>'
      str << '</c:numRef>'
      str << '</c:' << @name.to_s << '>'
    end

  end

end
