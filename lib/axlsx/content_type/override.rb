
# encoding: UTF-8
module Axlsx

  # An override content part. These parts are automatically created by for you based on the content of your package.
  class Override < AbstractContentType

    NODE_NAME = 'Override'

    # The name and location of the part.
    # @return [String]
    attr_reader :part_name
    alias :PartName :part_name

    # The name and location of the part.
    def part_name=(v) Axlsx::validate_string v; @part_name = v end
    alias :PartName= :part_name=

    def to_xml_string(str = '')
      super(NODE_NAME, str)
    end
  end
end
