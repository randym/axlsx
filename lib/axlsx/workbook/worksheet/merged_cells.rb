module Axlsx

  # A simple list of merged cells
  class MergedCells < SimpleTypedList

    # creates a new MergedCells object
    # @param [Worksheet] worksheet
    def initialize(worksheet)
      raise ArgumentError, 'you must provide a worksheet' unless worksheet.is_a?(Worksheet)
      super String
    end

    # adds cells to the merged cells collection
    # @param [Array||String] cells The cells to add to the merged cells
    # collection. This can be an array of actual cells or a string style
    # range like 'A1:C1'
    def add(cells)
      self << if cells.is_a?(String)
                 cells
               elsif cells.is_a?(Array)
                 Axlsx::cell_range(cells, false)
               elsif cells.is_a?(Row)
                 Axlsx::cell_range(cells, false)
               end
    end

    # serialize the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      return if empty?
      str << "<mergeCells count='#{size}'>"
      each { |merged_cell| str << "<mergeCell ref='#{merged_cell}'></mergeCell>" }
      str << '</mergeCells>'
    end
  end
end
