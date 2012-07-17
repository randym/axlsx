module Axlsx

  #A wraper class for comments that defines its on worksheet
  #serailization
  class WorksheetComments

    def initialize(worksheet)
      raise ArugumentError, 'You must provide a worksheet' unless worksheet.is_a?(Worksheet)
      @worksheet = worksheet
    end
    attr_reader :worksheet

    def comments
      @comments ||= Comments.new(worksheet)
    end

    def add_comment(options={})
      comments.add_comment(options)
    end 

    def relationships
      return [] unless has_comments?
      comments.relationships
    end

    def has_comments?
      !comments.empty?
    end

    def index
      worksheet.relationships.index { |r| r.Type == VML_DRAWING_R } + 1
    end
    def to_xml_string(str = '')
      return unless has_comments?
      str << "<legacyDrawing r:id='rId#{index}' />"
    end
  end
end
