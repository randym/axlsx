module Axlsx
  require 'axlsx/content_type/default.rb'
  require 'axlsx/content_type/override.rb'

  # ContentTypes used in the package. This is automatcially managed by the package package.
  class ContentType < SimpleTypedList
    
    def initialize
      super [Override, Default]
    end
    
    # Generates the xml document for [Content_Types].xml
    # @return [String] The document as a string.
    def to_xml()
      builder = Nokogiri::XML::Builder.new(:encoding => ENCODING) do |xml|
        xml.Types(:xmlns => Axlsx::XML_NS_T) {
          each { |type| type.to_xml(xml) }
        }
      end
      builder.to_xml
    end
  end
end
