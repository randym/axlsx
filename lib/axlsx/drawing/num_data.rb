# -*- coding: utf-8 -*-
module Axlsx

  #This class specifies data for a particular data point. It is used for both numCache and numLit object
  class NumData

    include Axlsx::OptionsParser

    # creates a new NumVal object
    # @option options [String] formatCode
    # @option options [Array] :data
    # @see StrData
    def initialize(options={})
      @format_code = "General"
      @pt = SimpleTypedList.new NumVal
      parse_options options
    end

    # A string representing the format code to apply. For more information see see the SpreadsheetML numFmt element's (§18.8.30) formatCode attribute.
    # @return [String]
    attr_reader :format_code

    # Creates the val objects for this data set. I am not overly confident this is going to play nicely with time and data types.
    # @param [Array] values An array of cells or values.
    def data=(values=[])
      @tag_name = values.first.is_a?(Cell) ? :numCache : :numLit
      values.each do |value|
        value = value.is_formula? ? 0 : value.value if value.is_a?(Cell)
        @pt << NumVal.new(:v => value)
      end
    end

    # @see format_code
    def format_code=(v='General')
      Axlsx::validate_string(v)
      @format_code = v
    end

    # serialize the object
    def to_xml_string(str = "")
      str << '<c:' << @tag_name.to_s << '>'
      str << '<c:formatCode>' << format_code.to_s << '</c:formatCode>'
      str << '<c:ptCount val="' << @pt.size.to_s << '"/>'
      @pt.each_with_index do |num_val, index|
        num_val.to_xml_string index, str
      end
      str << '</c:' << @tag_name.to_s << '>'
    end

  end

end
