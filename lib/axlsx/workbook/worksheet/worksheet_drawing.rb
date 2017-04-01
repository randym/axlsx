module Axlsx

  # This is a utility class for serialing the drawing node in a
  # worksheet. Drawing objects have their own serialization that exports
  # a drawing document. This is only for the single node in the
  # worksheet
  class WorksheetDrawing

    # Creates a new WorksheetDrawing
    # @param [Worksheet] worksheet
    def initialize(worksheet)
      raise ArgumentError, 'you must provide a worksheet' unless worksheet.is_a?(Worksheet)
      @worksheet = worksheet
      @drawing = nil
    end

    attr_reader :worksheet

    attr_reader :drawing

    # adds a chart to the drawing object
    # @param [Class] chart_type The type of chart to add
    # @param [Hash] options Options to pass on to the drawing and chart
    # @see Worksheet#add_chart
    def add_chart(chart_type, options)
      @drawing ||= Drawing.new worksheet
      drawing.add_chart(chart_type, options)
    end

    # adds an image to the drawing object
    # @param [Hash] options Options to pass on to the drawing and image
    # @see Worksheet#add_image
    def add_image(options)
      @drawing ||= Drawing.new(worksheet)
      drawing.add_image(options)
    end

    # helper method to tell us if the drawing has something in it or not
    # @return [Boolean]
    def has_drawing?
      @drawing.is_a? Drawing
    end

    # The relationship instance for this drawing.
    # @return [Relationship]
    def relationship
      return unless has_drawing?
      Relationship.new(self, DRAWING_R, "../#{drawing.pn}")
    end

    # Serialize the drawing for the worksheet
    # @param [String] str
    def to_xml_string(str = '')
      return unless has_drawing?
      str << "<drawing r:id='#{relationship.Id}'/>"
    end
  end
end
