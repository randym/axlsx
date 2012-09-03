module Axlsx
  class SignatureProperties < SimpleTypedList

    def initialize
      super SignatureProperty
    end

    def to_xml_string(str = '')
      str << '<SignatureProperties>'
      list.each { |item| item.to_xml_string(str) }
      str << '</SignatureProperties>'
    end
  end
end
