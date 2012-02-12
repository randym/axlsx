# encoding: UTF-8
module Axlsx
  # An default content part. These parts are automatically created by for you based on the content of your package.
  class Default

    # The extension of the content type.
    # @return [String]
    attr_reader :Extension

    # The type of content.
    # @return [String] 
    attr_reader :ContentType

    #Creates a new Default object
    # @option options [String] Extension
    # @option options [String] ContentType
    # @raise [ArgumentError] An argument error is raised if both Extension and ContentType are not specified.
    def initialize(options={})
      raise ArgumentError, "Extension and ContentType are required" unless options[:Extension] && options[:ContentType]
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end      
    end
    # Sets the file extension for this content type.
    def Extension=(v) Axlsx::validate_string v; @Extension = v end

    # Sets the content type
    # @see Axlsx#validate_content_type
    def ContentType=(v) Axlsx::validate_content_type v; @ContentType = v end

    # Serializes the object to xml
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml.Default(self.instance_values)
    end
  end
end
