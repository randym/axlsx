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

    # Creates a new TwoCellAnchor object
    # @param [Drawing] drawing
    # @param [Chart] chart
    # @option options [Array] start_at
    # @option options [Array] end_at
    def initialize(drawing, chart_type, options)
      @drawing = drawing
      drawing.anchors << self      

      @from, @to =  Marker.new, Marker.new(:col => 5, :row=>10)
      @graphic_frame = GraphicFrame.new(self, chart_type, options)

      self.start_at(options[:start_at][0], options[:start_at][1]) if options[:start_at].is_a?(Array)
      self.end_at(options[:end_at][0], options[:end_at][1]) if options[:end_at].is_a?(Array)
      # passing a reference to our start and end markers for convenience
      # this lets us access the markers directly from the chart.
      @graphic_frame.chart.send(:start_at=, @from)
      @graphic_frame.chart.send(:end_at=, @to)
    end

    def index
      @drawing.anchors.index(self)
    end
    
   
    # This is a short cut method to set the start anchor position
    # @param [Integer] x The column
    # @param [Integer] y The row
    # @return [Marker]
    def start_at(x, y)
      @from.col = x
      @from.row = y
      @from
    end

    # This is a short cut method to set the end anchor position
    # @param [Integer] x The column
    # @param [Integer] y The row
    # @return [Marker]
    def end_at(x, y)
      @to.col = x
      @to.row = y
      @to
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
  end
end
