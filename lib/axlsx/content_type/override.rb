module Axlsx
  # An override content part. These parts are automatically created by for you based on the content of your package.
  class Override

    # The type of content.
    # @return [String] 
    attr_reader :ContentType

    # The name and location of the part.
    # @return [String] 
    attr_reader :PartName

    #Creates a new Override object
    # @option options [String] PartName
    # @option options [String] ContentType
    # @raise [ArgumentError] An argument error is raised if both PartName and ContentType are not specified.
    def initialize(options={})
      raise ArgumentError, "PartName and ContentType are required" unless options[:PartName] && options[:ContentType]
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end      
    end

    # The name and location of the part.
    def PartName=(v) Axlsx::validate_string v; @PartName = v end

    # The content type. 
    # @see Axlsx#validate_content_type
    def ContentType=(v) Axlsx::validate_content_type v; @ContentType = v end

    # Serializes the Override object to xml
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    def to_xml(xml)
      xml.Override(self.instance_values)
    end
  end
end
