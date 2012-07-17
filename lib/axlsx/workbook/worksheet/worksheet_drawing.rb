module Axlsx
  
  # This is a utility class for serialing the drawing node in a
  # worksheet. Drawing objects have their own serialization that exports
  # a drawing document. This is only for the single node in the
  # worksheet
  class WorksheetDrawing

    def initialize(worksheet)
      raise ArgumentError, 'you must provide a worksheet' unless worksheet.is_a?(Worksheet)
      @worksheet = worksheet      
      @drawing = nil
    end

    attr_reader :worksheet
    attr_reader :drawing

    def add_chart(chart_type, options)
      @drawing ||= Drawing.new worksheet
      drawing.add_chart(chart_type, options)
    end
   
    def add_image(options)
      @drawing ||= Drawing.new worksheet
      drawing.add_image(options)
    end 
   
    def has_drawing?
      @drawing.is_a? Drawing
    end

    # The relationship required by this object
    # @return [Relationship]
    def relationship
      return unless has_drawing?
      Relationship.new(DRAWING_R, "../#{drawing.pn}") 
    end

    def index
       worksheet.relationships.index{ |r| r.Type == DRAWING_R } +1 
    end

    def to_xml_string(str = '')
      return unless has_drawing? 
      str << "<drawing r:id='rId#{index}'/>"
    end
  end
end
