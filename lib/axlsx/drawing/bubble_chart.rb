# encoding: UTF-8
module Axlsx

  # The BubbleChart allows you to insert a bubble chart into your worksheet
  # @see Worksheet#add_chart
  # @see Chart#add_series
  # @see README for an example
  class BubbleChart < Chart

    include Axlsx::OptionsParser

    # the x value axis
    # @return [ValAxis]
    def x_val_axis
      axes[:x_val_axis]
    end
    alias :xValAxis :x_val_axis

    # the y value axis
    # @return [ValAxis]
    def y_val_axis
      axes[:y_val_axis]
    end
    alias :yValAxis :y_val_axis

    # Creates a new bubble chart
    def initialize(frame, options={})
      @vary_colors = 0

           super(frame, options)
      @series_type = BubbleSeries
      @d_lbls = nil
      parse_options options
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      super(str) do
        str << '<c:bubbleChart>'
        str << ('<c:varyColors val="' << vary_colors.to_s << '"/>')
        @series.each { |ser| ser.to_xml_string(str) }
        d_lbls.to_xml_string(str) if @d_lbls
        axes.to_xml_string(str, :ids => true)
        str << '</c:bubbleChart>'
        axes.to_xml_string(str)
      end
      str
    end

    # The axes for the bubble chart. BubbleChart has an x_val_axis and
    # a y_val_axis
    # @return [Axes]
    def axes
      @axes ||= Axes.new(:x_val_axis => ValAxis, :y_val_axis => ValAxis)
    end
  end
end
