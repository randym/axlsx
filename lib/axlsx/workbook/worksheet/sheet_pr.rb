module Axlsx
  class SheetPr

    def initialize(worksheet)
      raise ArgumentError, "you must provide a worksheet" unless worksheet.is_a?(Worksheet)
      @worksheet = worksheet
    end

    attr_reader :worksheet

    def to_xml_string(str = '')
       return unless worksheet.fit_to_page?
      str << "<sheetPr><pageSetUpPr fitToPage=\"%s\"></pageSetUpPr></sheetPr>" % worksheet.fit_to_page?
    end
  end
end
