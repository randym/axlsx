module Axlsx

  # This class manages the dimensions for a worksheet.
  # While this node is optional in the specification some readers like 
  # LibraOffice require this node to render the sheet
  class Dimension

    # the default value for the first cell in the dimension
    # @return [String]
    def self.default_first
      @@default_first ||= 'A1'
    end

    # the default value for the last cell in the dimension
    # @return [String]
    def self.default_last
      @@default_last ||= 'AA200'
    end

    # Creates a new dimension object
    # @param[Worksheet] worksheet - the worksheet this dimension applies
    # to.
    def initialize(worksheet)
      raise ArgumentError, "you must provide a worksheet" unless worksheet.is_a?(Worksheet)
      @worksheet = worksheet
    end

    attr_reader :worksheet

    # the full refernece for this dimension
    # @return [String]
    def sqref
      "#{first_cell_reference}:#{last_cell_reference}"
    end 

    # serialize the object
    # @return [String]
    def to_xml_string(str = '')
      return if worksheet.rows.empty?
      str << "<dimension ref=\"%s\"></dimension>" % sqref
    end

    # The first cell in the dimension
    # @return [String]
    def first_cell_reference
      dimension_reference(worksheet.rows.first.first, Dimension.default_first)
    end

    # the last cell in the dimension
    # @return [String]
    def last_cell_reference
      dimension_reference(worksheet.rows.last.last, Dimension.default_last)
    end

    private

    # Returns the reference of a cell or the default specified
    # @return [String]
    def dimension_reference(cell, default)
      return default unless cell.respond_to?(:r)
      cell.r
    end
  end
end
