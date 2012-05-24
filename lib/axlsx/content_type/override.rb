# encoding: UTF-8
module Axlsx
  # An override content part. These parts are automatically created by for you based on the content of your package.
  class Override

    # The type of content.
    # @return [String]
    attr_reader :content_type
    alias :ContentType :content_type

    # The name and location of the part.
    # @return [String]
    attr_reader :part_name
    alias :PartName :part_name

    #Creates a new Override object
    # @option options [String] PartName
    # @option options [String] ContentType
    # @raise [ArgumentError] An argument error is raised if both PartName and ContentType are not specified.
    def initialize(options={})
      raise ArgumentError, "PartName and ContentType are required" unless (options[:PartName] || options[:part_name]) && (options[:ContentType] || options[:content_type])
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
    end

    # The name and location of the part.
    def part_name=(v) Axlsx::validate_string v; @part_name = v end
    alias :PartName= :part_name=

    # The content type.
    # @see Axlsx#validate_content_type
    def content_type=(v) Axlsx::validate_content_type v; @content_type = v end
    alias :ContentType= :content_type=

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<Override '
      str << instance_values.map { |key, value| '' << Axlsx::camel(key) << '="' << value.to_s << '"' }.join(' ')
      str << '/>'
    end

  end
end
