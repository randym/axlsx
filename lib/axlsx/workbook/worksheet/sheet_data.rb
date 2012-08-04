module Axlsx

  # This class manages the serialization of rows for worksheets
  class SheetData

    # Creates a new SheetData object
    # @param [Worksheet] worksheet The worksheet that owns this sheet data.
    def initialize(worksheet)
      raise ArgumentError, "you must provide a worksheet" unless worksheet.is_a?(Worksheet)
      @worksheet = worksheet
    end
    
    attr_reader :worksheet

    # Serialize the sheet data
    # @param [String] str the string this objects serializaton will be concacted to.
    # @return [String]
    def to_xml_string(str = '')
      str << '<sheetData>'
      worksheet.rows.each_with_index{ |row, index| row.to_xml_string(index, str) }
      str << '</sheetData>'
    end
    
  end
end
