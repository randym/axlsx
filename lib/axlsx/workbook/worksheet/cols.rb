module Axlsx

  # The cols class manages the col object used to manage column widths. 
  # This is where the magic happens with autowidth
  class Cols < SimpleTypedList

    def initialize(worksheet)
      raise ArgumentError, "you must provide a worksheet" unless worksheet.is_a?(Worksheet)
      super Col
      @worksheet = worksheet
    end

    def to_xml_string(str = '')
     return if empty?
     str << '<cols>'
     each { |item| item.to_xml_string(str) }
     str << '</cols>' 
    end
  end
end
