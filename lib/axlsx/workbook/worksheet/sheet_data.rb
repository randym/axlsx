module Axlsx

  # This class manages the serialization of rows for worksheets
  class SheetData

    def initialize(worksheet)
      raise ArgumentError, "you must provide a worksheet" unless worksheet.is_a?(Worksheet)
      @worksheet = worksheet
    end
    
    attr_reader :worksheet

    def to_xml_string(str = '')
      str << '<sheetData>'
      worksheet.rows.each_with_index{ |row, index| row.to_xml_string(index, str) }
      str << '</sheetData>'
    end
    
  end
end
