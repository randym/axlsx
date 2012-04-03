# encoding: UTF-8
module Axlsx
  # CellProtection stores information about locking or hiding cells in spreadsheet.
  # @note Using Styles#add_style is the recommended way to manage cell protection.
  # @see Styles#add_style
  class CellProtection

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
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
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
      str << instance_values.map { |key, value| '' << key.to_s << '="' << value.to_s << '"' }.join(' ')
      str << '/>'
    end

  end
end
