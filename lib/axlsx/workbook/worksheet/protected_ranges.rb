module Axlsx

  # A self serializing collection of ranges that should be protected in
  # the worksheet
  class ProtectedRanges < SimpleTypedList

    attr_reader :worksheet

    def initialize(worksheet)
      raise ArgumentError, 'You must provide a worksheet' unless worksheet.is_a?(Worksheet)
      super ProtectedRange
      @worksheet = worksheet
    end

    # Adds a protected range
    # @param [Array|String] cells A string range reference or array of cells that will be protected
    def add_range(cells)
     sqref = if cells.is_a?(String)
               cells
             elsif cells.is_a?(SimpleTypedList) || cells.is_a?(Array)
               Axlsx::cell_range(cells, false)
             end
     self << ProtectedRange.new(:sqref => sqref, :name => "Range#{size}")
     last
    end

    # Serializes the protected ranges
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      return if empty?
      str << '<protectedRanges>'
      each { |range| range.to_xml_string(str) }
      str << '</protectedRanges>'
    end
  end
end
