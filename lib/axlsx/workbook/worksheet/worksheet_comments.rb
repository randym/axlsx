module Axlsx

  # A wraper class for comments that defines its on worksheet
  # serailization
  class WorksheetComments

    # Creates a new WorksheetComments object
    # param [Worksheet] worksheet The worksheet comments in thes object belong to
    def initialize(worksheet)
      raise ArugumentError, 'You must provide a worksheet' unless worksheet.is_a?(Worksheet)
      @worksheet = worksheet
    end

    attr_reader :worksheet

    # The comments for this worksheet.
    # @return [Comments]
    def comments
      @comments ||= Comments.new(worksheet)
    end

    # Adds a comment
    # @param [Hash] options
    # @see Comments#add_comment
    def add_comment(options={})
      comments.add_comment(options)
    end 

    # The relationships defined by this objects comments collection
    # @return [Relationships]
    def relationships
      return [] unless has_comments?
      comments.relationships
    end


    # Helper method to tell us if there are comments in the comments collection
    # @return [Boolean]
    def has_comments?
      !comments.empty?
    end

    # The relationship id of the VML drawing that will render the comments.
    # @see Relationship#Id
    # @return [String]
    def drawing_rId
      comments.relationships.find{ |r| r.Type == VML_DRAWING_R }.Id
    end

    # Seraalize the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      return unless has_comments?
      str << "<legacyDrawing r:id='#{drawing_rId}' />"
    end
  end
end
