# encoding: UTF-8
module Axlsx
  # This class details the anchor points for drawings.
  # @note The recommended way to manage drawings and charts is Worksheet#add_chart. Anchors are specified by the :start_at and :end_at options to that method.
  # @see Worksheet#add_chart
  class TwoCellAnchor

    include Axlsx::OptionsParser

    # A marker that defines the from cell anchor. The default from column and row are 0 and 0 respectively
    # @return [Marker]
    attr_reader :from
    # A marker that returns the to cell anchor. The default to column and row are 5 and 10 respectively
    # @return [Marker]
    attr_reader :to

    # The frame for your chart
    # @note this will be discontinued in version 2.0 please use object
    # @return [GraphicFrame]
    # attr_reader :graphic_frame

    # The object this anchor hosts
    # @return [Pic, GraphicFrame]
    attr_reader :object

    # The drawing that holds this anchor
    # @return [Drawing]
    attr_reader :drawing

    # Creates a new TwoCellAnchor object
    # c.start_at 5, 9
    # @param [Drawing] drawing
    # @option options [Array] :start_at the col, row to start at THIS IS DOCUMENTED BUT NOT IMPLEMENTED HERE!
    # @option options [Array] :end_at the col, row to end at
    def initialize(drawing, options={})
      @drawing = drawing
      drawing.anchors << self
      @from, @to =  Marker.new, Marker.new(:col => 5, :row=>10)
      parse_options options

      # bit of a hack to work around the fact that the coords for start at and end at
      # are passed in as an array when specified in intialization options - however
      start_at(*options[:start_at]) if options[:start_at]
      end_at(*options[:end_at]) if options[:end_at]
    end

    # sets the col, row attributes for the from marker.
    # @note The recommended way to set the start position for graphical
    # objects is directly thru the object.
    # @see Chart#start_at
    def start_at(x, y=nil)
      from.coord x, y
    end

    # sets the col, row attributes for the to marker
    # @note the recommended way to set the to position for graphical
    # objects is directly thru the object
    # @see Char#end_at
    def end_at(x, y=nil)
      to.coord x, y
    end

    # Creates a graphic frame and chart object associated with this anchor
    # @return [Chart]
    def add_chart(chart_type, options)
      @object = GraphicFrame.new(self, chart_type, options)
      @object.chart
    end

    # Creates an image associated with this anchor.
    def add_pic(options={})
      @object = Pic.new(self, options)
    end

    # The index of this anchor in the drawing
    # @return [Integer]
    def index
      @drawing.anchors.index(self)
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<xdr:twoCellAnchor>'
      str << '<xdr:from>'
      from.to_xml_string str
      str << '</xdr:from>'
      str << '<xdr:to>'
      to.to_xml_string str
      str << '</xdr:to>'
      object.to_xml_string(str)
      str << '<xdr:clientData/>'
      str << '</xdr:twoCellAnchor>'
    end
  end
end
