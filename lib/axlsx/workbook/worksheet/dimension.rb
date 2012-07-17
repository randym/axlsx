module Axlsx

  # This class manages the dimensions for a worksheet.
  # While this node is optional in the specification some readers like 
  # LibraOffice require this node to render the sheet
  class Dimension


    def self.default_first
      @@default_first ||= 'A1'
    end

    def self.default_last
      @@default_last ||= 'AA200'
    end


    def initialize(worksheet)
      raise ArgumentError, "you must provide a worksheet" unless worksheet.is_a?(Worksheet)
      @worksheet = worksheet
    end

    attr_reader :worksheet

    def sqref
      "#{first_cell_reference}:#{last_cell_reference}"
    end 

    def to_xml_string(str = '')
      return if worksheet.rows.empty?
      str << "<dimension ref=\"%s\"></dimension>" % sqref
    end

    def first_cell_reference
      dimension_reference(worksheet.rows.first.cells.first, Dimension.default_first)
    end

    def last_cell_reference
      dimension_reference(worksheet.rows.last.cells.last, Dimension.default_last)
    end

    private

    def dimension_reference(cell, default)
      return default unless cell.respond_to?(:r)
      cell.r
    end
  end
end
