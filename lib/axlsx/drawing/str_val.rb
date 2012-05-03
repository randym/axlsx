# -*- coding: utf-8 -*-
module Axlsx

  #This class specifies data for a particular data point.
  class StrVal

    # a string value.
    # @return [String]
    attr_reader :v

    # creates a new StrVal object
    # @option options [String] v
    def initialize(options={})
      @v = ""
      @idx = 0
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
    end
    # @see v
    def v=(v)
      @v = v.to_s
    end

    # serialize the object
    def to_xml_string(idx, str = "")
      Axlsx::validate_unsigned_int(idx)
      str << '<c:pt idx="' << idx.to_s << '"><c:v>' << v.to_s << '</c:v></c:pt>'
    end

  end

end
