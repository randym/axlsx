# encoding: UTF-8
module Axlsx
  require 'axlsx/content_type/abstract_content_type.rb'
  require 'axlsx/content_type/default.rb'
  require 'axlsx/content_type/override.rb'

  # ContentTypes used in the package. This is automatically managed by the package package.
  class ContentType < SimpleTypedList

    def initialize
      super [Override, Default]
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<?xml version="1.0" encoding="UTF-8"?>'
      str << '<Types xmlns="' << XML_NS_T << '">'
      each { |type| type.to_xml_string(str) }
      str << '</Types>'
    end

  end

end
