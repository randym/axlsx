# encoding: UTF-8
module Axlsx

  # An override content part. These parts are automatically created by for you based on the content of your package.
  class Override

    #Creates a new Override object
    # @option options [String] PartName
    # @option options [String] ContentType
    # @raise [ArgumentError] An argument error is raised if both PartName and ContentType are not specified.
    def initialize(options={})
      raise ArgumentError, INVALID_ARGUMENTS unless validate_options(options)
      options.each do |name, value|
        self.send("#{name}=", value) if self.respond_to? "#{name}="
      end
    end

    # Error message for invalid options
    INVALID_ARGUMENTS = 'part_name and content_type are required'

    # The type of content.
    # @return [String]
    attr_reader :content_type
    alias :ContentType :content_type

    # The name and location of the part.
    # @return [String]
    attr_reader :part_name
    alias :PartName :part_name

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

    private 

    def validate_options(options)
      (options[:PartName] || options[:part_name]) && (options[:ContentType] || options[:content_type])
    end

  end

end
