# encoding: UTF-8
module Axlsx
  # an element of style that belongs to a table style.
  # @note tables and table styles are not supported in this version. This class exists in preparation for that support.
  class TableStyleElement

    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes

    # creates a new TableStyleElement object
    # @option options [Symbol] type
    # @option options [Integer] size
    # @option options [Integer] dxfId
    def initialize(options={})
      parse_options options
    end

    serializable_attributes :type, :size, :dxfId

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
      serialized_tag('tableStyleElement', str)
    end

  end
end
