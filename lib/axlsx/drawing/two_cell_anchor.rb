module Axlsx
  # This class details the anchor points for drawings.
  # @note The recommended way to manage drawings and charts is Worksheet#add_chart. Anchors are specified by the :start_at and :end_at options to that method.
  # @see Worksheet#add_chart
  class TwoCellAnchor

    # A marker that defines the from cell anchor. The default from column and row are 0 and 0 respectively
    # @return [Marker]
    attr_reader :from
    # A marker that returns the to cell anchor. The default to column and row are 5 and 10 respectively
    # @return [Marker]
    attr_reader :to

    # The frame for your chart
    # @return [GraphicFrame]
    attr_reader :graphic_frame

    # The drawing that holds this anchor
    # @return [Drawing]
    attr_reader :drawing

    # The index of this anchor in the drawing
    # @return [Integer]
    attr_reader :index

    # Creates a new TwoCellAnchor object and sets up a reference to the from and to markers in the 
    # graphic_frame's chart. That means that you can do stuff like
    # c = worksheet.add_chart Axlsx::Chart
    # c.start_at 5, 9
    # @param [Drawing] drawing
    # @param [Class] chart_type This is passed to the graphic frame for instantiation. must be Chart or a subclass of Chart
    # @option options [Array] start_at the col, row to start at
    # @option options [Array] end_at the col, row to end at
    def initialize(drawing, chart_type, options)
      @drawing = drawing
      drawing.anchors << self      
      @from, @to =  Marker.new, Marker.new(:col => 5, :row=>10)
      @graphic_frame = GraphicFrame.new(self, chart_type, options)
    end

    def index
      @drawing.anchors.index(self)
    end
    # Serializes the two cell anchor
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      #build it for now, break it down later!
      xml.send('xdr:twoCellAnchor') {
        xml.send('xdr:from') {
          from.to_xml(xml)
        }
        xml.send('xdr:to') {
          to.to_xml(xml)
        }
        @graphic_frame.to_xml(xml)
        xml.send('xdr:clientData')
      }
    end    

    private

  end
end
