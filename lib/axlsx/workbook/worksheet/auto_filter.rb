module Axlsx

  #This class represents an auto filter range in a worksheet
  class AutoFilter

    # creates a new Autofilter object
    # @param [Worksheet] worksheet
    def initialize(worksheet)
      raise ArgumentError, 'you must provide a worksheet' unless worksheet.is_a?(Worksheet)
      @worksheet = worksheet
    end

    attr_reader :worksheet

    # The range the autofilter should be applied to.
    # This should be a string like 'A1:B8'
    # @return [String]
    attr_accessor :range

    # the formula for the defined name required for this auto filter
    # @return [String]
    def defined_name
      return unless range
      Axlsx.cell_range(range.split(':').collect { |name| worksheet.name_to_cell(name)})
    end

    # serialize the object
    # @return [String]
    def to_xml_string(str='')
      str << "<autoFilter ref='#{range}'></autoFilter>"
    end

  end
end
