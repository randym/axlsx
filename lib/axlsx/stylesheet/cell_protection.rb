# encoding: UTF-8
module Axlsx
  # CellProtection stores information about locking or hiding cells in spreadsheet.
  # @note Using Styles#add_style is the recommended way to manage cell protection.
  # @see Styles#add_style
  class CellProtection

    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes

    serializable_attributes :hidden, :locked

    # specifies locking for cells that have the style containing this protection
    # @return [Boolean]
    attr_reader :hidden

    # specifies if the cells that have the style containing this protection
    # @return [Boolean]
    attr_reader :locked

    # Creates a new CellProtection
    # @option options [Boolean] hidden value for hidden protection
    # @option options [Boolean] locked value for locked protection
    def initialize(options={})
      parse_options options
    end

    # @see hidden
    def hidden=(v) Axlsx::validate_boolean v; @hidden = v end
    # @see locked
    def locked=(v) Axlsx::validate_boolean v; @locked = v end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<protection '
      serialized_attributes str
      str << '/>'
    end

  end
end
