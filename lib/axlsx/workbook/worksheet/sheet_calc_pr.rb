module Axlsx

  class SheetCalcPr

    def initialize

    end

    def to_xml_string(str='')
      str << '<sheetCalcPr fullCalcOnLoad="1"/>'
    end
  end
end
