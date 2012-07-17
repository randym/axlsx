module Axlsx

  class MergedCells < SimpleTypedList

    def initialize(worksheet)
      raise ArgumentError, 'you must provide a worksheet' unless worksheet.is_a?(Worksheet)
      super String
    end

    def add(cells)
      @list << if cells.is_a?(String)
                 cells
               elsif cells.is_a?(Array)
                 Axlsx::cell_range(cells, false)
               end
    end

    def to_xml_string(str = '')
      return if @list.empty?
      str << "<mergeCells count='#{size}'>"
      each { |merged_cell| str << "<mergeCell ref='#{merged_cell}'></mergeCell>" }
      str << '</mergeCells>'
    end
  end
end
