# encoding: UTF-8
module Axlsx
  # The Parser module mixes in a number of methods to help in generating a model from xml
  # This module is not included in the axlsx library at this time. It is for future development only,
  module Parser

    # The xml to be parsed
    attr_accessor :parser_xml
    
    # parse and assign string attribute
    def parse_string attr_name, xpath
      send("#{attr_name.to_s}=", parse_value(xpath))
    end
    
    # parse convert and assign node text to symbol
    def parse_symbol attr_name, xpath
      v = parse_value xpath
      v = v.to_sym unless v.nil?
      send("#{attr_name.to_s}=", v)
    end
    
    # parse, convert and assign note text to integer
    def parse_integer attr_name, xpath
      v = parse_value xpath
      v = v.to_i if v.respond_to?(:to_i)
      send("#{attr_name.to_s}=", v)
    end
    
    # parse, convert and assign node text to float
    def parse_float attr_name, xpath
      v = parse_value xpath
      v = v.to_f if v.respond_to?(:to_f)
      send("#{attr_name.to_s}=", v)
    end

    # return node text based on xpath
    def parse_value xpath
      node = parser_xml.xpath(xpath)
      return nil if node.empty?
      node.text.strip
    end    
    
  end
end
