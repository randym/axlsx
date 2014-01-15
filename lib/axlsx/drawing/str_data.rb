# -*- coding: utf-8 -*-
module Axlsx

  #This specifies the last string data used for a chart. (e.g. strLit and strCache)
  # This class is extended for NumData to include the formatCode attribute required for numLit and numCache
  class StrData

    include Axlsx::OptionsParser

    # creates a new StrVal object
    # @option options [Array] :data
    # @option options [String] :tag_name
    def initialize(options={})
      @tag_prefix = :str
      @type = StrVal
      @pt = SimpleTypedList.new(@type)
      parse_options options
    end

    # Creates the val objects for this data set. I am not overly confident this is going to play nicely with time and data types.
    # @param [Array] values An array of cells or values.
    def data=(values=[])
      @tag_name = values.first.is_a?(Cell) ? :strCache : :strLit
      values.each do |value|
        v = value.is_a?(Cell) ? value.value : value
        @pt << @type.new(:v => v)
      end
    end

    # serialize the object
    def to_xml_string(str = "")
      str << ('<c:' << @tag_name.to_s << '>')
      str << ('<c:ptCount val="' << @pt.size.to_s << '"/>')
      @pt.each_with_index do |value, index|
        value.to_xml_string index, str
      end
      str << ('</c:' << @tag_name.to_s << '>')
    end

  end

end
