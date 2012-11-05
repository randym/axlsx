# encoding: UTF-8
module Axlsx

  # An default content part. These parts are automatically created by for you based on the content of your package.
  class Default < AbstractContentType

    # The serialization node name for this class
    NODE_NAME = 'Default'

    # The extension of the content type.
    # @return [String]
    attr_reader :extension
    alias :Extension :extension

    # Sets the file extension for this content type.
    def extension=(v) Axlsx::validate_string v; @extension = v end
    alias :Extension= :extension=

    # Serializes this object to xml
    def to_xml_string(str ='')
      super(NODE_NAME, str)
    end
  end

end
