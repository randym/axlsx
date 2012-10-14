module Axlsx

  # the SheetCalcPr object for the worksheet
  # This object contains calculation properties for the worksheet.
  class SheetCalcPr
   include Axlsx::OptionsParser
   include Axlsx::SerializedAttributes
   include Axlsx::Accessors
    # creates a new SheetCalcPr
    # @param [Hash] options Options for this object
    # @option [Boolean] full_calc_on_load @see full_calc_on_load
    def initialize(options={})
      @full_calc_on_load = true
      parse_options options
    end

    boolean_attr_accessor :full_calc_on_load

    serializable_attributes :full_calc_on_load

    # Serialize the object
    # @param [String] str the string to append this objects serialized
    # content to.
    # @return [String]
    def to_xml_string(str='')
      str << "<sheetCalcPr #{serialized_attributes}/>"
    end
  end
end
