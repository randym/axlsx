# encoding: UTF-8
module Axlsx

  # The ScatterChart allows you to insert a scatter chart into your worksheet
  # @see Worksheet#add_chart
  # @see Chart#add_series
  # @see README for an example
  class ScatterChart < Chart

    include Axlsx::OptionsParser

    # The Style for the scatter chart
    # must be one of :none | :line | :lineMarker | :marker | :smooth | :smoothMarker
    # return [Symbol]
    attr_reader :scatter_style
    alias :scatterStyle :scatter_style

    # the x value axis
    # @return [ValAxis]
    def x_val_axis
      axes[:x_val_axis]
    end
    alias :xValAxis :x_val_axis

    # the y value axis
    # @return [ValAxis]
    def y_val_axis
      axes[:x_val_axis]
    end
    alias :yValAxis :y_val_axis

    # Creates a new scatter chart
    def initialize(frame, options={})
      @vary_colors = 0
      @scatter_style = :lineMarker

           super(frame, options)
      @series_type = ScatterSeries
      @d_lbls = nil
      parse_options options
    end

    # see #scatterStyle
    def scatter_style=(v)
      Axlsx.validate_scatter_style(v)
      @scatter_style = v
    end
    alias :scatterStyle= :scatter_style=

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      super(str) do |str_inner|
        str_inner << '<c:scatterChart>'
        str_inner << '<c:scatterStyle val="' << scatter_style.to_s << '"/>'
        str_inner << '<c:varyColors val="' << vary_colors.to_s << '"/>'
        @series.each { |ser| ser.to_xml_string(str_inner) }
        d_lbls.to_xml_string(str_inner) if @d_lbls
        axes.to_xml_string(str_inner, :ids => true)
        str_inner << '</c:scatterChart>'
        axes.to_xml_string(str_inner)
      end
      str
    end

    # The axes for the scatter chart. ScatterChart has an x_val_axis and
    # a y_val_axis
    # @return [Axes]
    def axes
      @axes ||= Axes.new(:x_val_axis => ValAxis, :y_val_axis => ValAxis)
    end
  end
end
