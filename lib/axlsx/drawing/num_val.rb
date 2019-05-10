# -*- coding: utf-8 -*-
# frozen_string_literal: true
module Axlsx

  #This class specifies data for a particular data point.
  class NumVal < StrVal

    # A string representing the format code to apply.
    # For more information see see the SpreadsheetML numFmt element's (§18.8.30) formatCode attribute.
    # @return [String]
    attr_reader :format_code

    # creates a new NumVal object
    # @option options [String] formatCode
    # @option options [Integer] v
    def initialize(options={})
      @format_code = "General"
      super(options)
    end

    # @see format_code
    def format_code=(v)
      Axlsx::validate_string(v)
      @format_code = v
    end

    # serialize the object
    def to_xml_string(idx, str = String.new)
      Axlsx::validate_unsigned_int(idx)
      if !v.to_s.empty?
        str << "<c:pt idx=\"#{idx}\" formatCode=\"#{format_code}\"><c:v>#{v}</c:v></c:pt>"
      end
    end
  end
end
