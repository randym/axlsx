module Axlsx

  # This class extracts the common parts from Default and Override
  class AbstractContentType

    include Axlsx::OptionsParser

    # Initializes an abstract content type
    # @see Default, Override
    def initialize(options={})
       parse_options options
    end

    # The type of content.
    # @return [String]
    attr_reader :content_type
    alias :ContentType :content_type

    # The content type.
    # @see Axlsx#validate_content_type
    def content_type=(v) Axlsx::validate_content_type v; @content_type = v end
    alias :ContentType= :content_type=

    # Serialize the contenty type to xml
    def to_xml_string(node_name = '', str = '')
      str << "<#{node_name} "
      str << instance_values.map { |key, value| Axlsx::camel(key) << '="' << value.to_s << '"' }.join(' ')
      str << '/>'
    end

  end
end
