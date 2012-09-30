module Axlsx

  # the SheetCalcPr object for the worksheet
  # This object contains calculation properties for the worksheet.
  class SheetCalcPr

    # creates a new SheetCalcPr
    # @param [Hash] options Options for this object
    # @option [Boolean] full_calc_on_load @see full_calc_on_load
    def initialize(options={})
      @full_calc_on_load = true
      options.each do |key, value|
        self.send("#{key}=", value) if self.respond_to?("#{key}=")
      end
    end

    # Indicates whether the application should do a full calculate on
    # load due to contents on ￼this sheet. After load and successful cal,c
    # the application shall set this value to false. Set ￼this to true
    # when the application should calculate the workbook on load.
    # @return [Boolean]
    def full_calc_on_load
      @full_calc_on_load
    end

    # specify the full_calc_on_load value
    # @param [Boolean] value
    # @see full_calc_on_load
    def full_calc_on_load=(value)
      Axlsx.validate_boolean(value)
      @full_calc_on_load = value
    end

    # Serialize the object
    # @param [String] str the string to append this objects serialized
    # content to.
    # @return [String]
    def to_xml_string(str='')
      str << "<sheetCalcPr #{serialized_attributes}/>"
    end

    private

    def serialized_attributes
      instance_values.map { |key, value| "#{Axlsx.camel(key, false)}='#{value}'" }.join(' ')
    end

  end
end
