# encoding: UTF-8
module Axlsx
  # an element of style that belongs to a table style.
  # @note tables and table styles are not supported in this version. This class exists in preparation for that support.
  class TableStyleElement
    # The type of style element. The following type are allowed
    #   :wholeTable
    #   :headerRow
    #   :totalRow
    #   :firstColumn
    #   :lastColumn
    #   :firstRowStripe
    #   :secondRowStripe
    #   :firstColumnStripe
    #   :secondColumnStripe
    #   :firstHeaderCell
    #   :lastHeaderCell
    #   :firstTotalCell
    #   :lastTotalCell
    #   :firstSubtotalColumn
    #   :secondSubtotalColumn
    #   :thirdSubtotalColumn
    #   :firstSubtotalRow
    #   :secondSubtotalRow
    #   :thirdSubtotalRow
    #   :blankRow
    #   :firstColumnSubheading
    #   :secondColumnSubheading
    #   :thirdColumnSubheading
    #   :firstRowSubheading
    #   :secondRowSubheading
    #   :thirdRowSubheading
    #   :pageFieldLabels
    #   :pageFieldValues
    # @return [Symbol]
    attr_reader :type

    # Number of rows or columns used in striping when the type is firstRowStripe, secondRowStripe, firstColumnStripe, or secondColumnStripe.
    # @return [Integer]
    attr_reader :size

    # The dxfId this style element points to
    # @return [Integer]
    attr_reader :dxfId

    # creates a new TableStyleElement object
    # @option options [Symbol] type
    # @option options [Integer] size
    # @option options [Integer] dxfId
    def initialize(options={})
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? o[0]
      end
    end

    # @see type
    def type=(v) Axlsx::validate_table_element_type v; @type = v end

    # @see size
    def size=(v) Axlsx::validate_unsigned_int v; @size = v end

    # @see dxfId
    def dxfId=(v) Axlsx::validate_unsigned_int v; @dxfId = v end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<tableStyleElement '
      str << instance_values.map { |key, value| '' << key.to_s << '="' << value.to_s << '"' }.join(' ')
      str << '/>'
    end

  end
end
