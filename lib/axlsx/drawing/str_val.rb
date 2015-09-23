# -*- coding: utf-8 -*-
module Axlsx

  #This class specifies data for a particular data point.
  class StrVal

    include Axlsx::OptionsParser

    # creates a new StrVal object
    # @option options [String] v
    def initialize(options={})
      @v = ""
      @idx = 0
      parse_options options
    end

    # a string value.
    # @return [String]
    attr_reader :v

    # @see v
    def v=(v)
      @v = v.to_s
    end

    # serialize the object
    def to_xml_string(idx, str = "")
      Axlsx::validate_unsigned_int(idx)
      if !v.to_s.empty?
        str << ('<c:pt idx="' << idx.to_s << '"><c:v>' << ::CGI.escapeHTML(v.to_s) << '</c:v></c:pt>')
      end
    end
  end
end
