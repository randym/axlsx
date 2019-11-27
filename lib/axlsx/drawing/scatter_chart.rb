# encoding: UTF-8
# frozen_string_literal: true
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
      axes[:y_val_axis]
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
    def to_xml_string(str = String.new)
      super(str) do
        str << '<c:scatterChart>'\
               "<c:scatterStyle val=\"#{scatter_style}\"/>"\
               "<c:varyColors val=\"#{vary_colors}\"/>"
        @series.each { |ser| ser.to_xml_string(str) }
        d_lbls.to_xml_string(str) if @d_lbls
        axes.to_xml_string(str, :ids => true)
        str << '</c:scatterChart>'
        axes.to_xml_string(str)
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
