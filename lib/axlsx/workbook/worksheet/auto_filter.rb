module Axlsx

  #This class represents an auto filter range in a worksheet
  class AutoFilter
    def initialize(worksheet)
      raise ArgumentError, 'you must provide a worksheet' unless worksheet.is_a?(Worksheet)
      @worksheet = worksheet
    end
    
    attr_reader :worksheet
    attr_accessor :range

    def defined_name
      return unless range
      Axlsx.cell_range(range.split(':').collect { |name| worksheet.name_to_cell(name)})
                            end

    def to_xml_string(str='')
      str << "<autoFilter ref='#{range}'></autoFilter>"
    end
    
  end
end
