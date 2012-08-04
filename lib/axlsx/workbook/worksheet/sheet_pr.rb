module Axlsx

  # The SheetPr class manages serialization fo a worksheet's sheetPr element.
  # Only fit_to_page is implemented
  class SheetPr

    # Creates a new SheetPr object
    # @param [Worksheet] worksheet The worksheet that owns this SheetPr object
    def initialize(worksheet)
      raise ArgumentError, "you must provide a worksheet" unless worksheet.is_a?(Worksheet)
      @worksheet = worksheet
    end

    attr_reader :worksheet

    # Serialize the object
    # @param [String] str serialized output will be appended to this object if provided.
    # @return [String]
    def to_xml_string(str = '')
       return unless worksheet.fit_to_page?
      str << "<sheetPr><pageSetUpPr fitToPage=\"%s\"></pageSetUpPr></sheetPr>" % worksheet.fit_to_page?
    end
  end
end
