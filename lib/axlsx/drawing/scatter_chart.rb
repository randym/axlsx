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
    attr_reader :scatterStyle

    # the x value axis
    # @return [ValAxis]
    attr_reader :xValAxis

    # the y value axis
    # @return [ValAxis]
    attr_reader :yValAxis

    # Creates a new scatter chart
    def initialize(frame, options={})
      @vary_colors = 0
      @scatterStyle = :lineMarker
      @xValAxId = rand(8 ** 8)
      @yValAxId = rand(8 ** 8)
      @xValAxis = ValAxis.new(@xValAxId, @yValAxId)
      @yValAxis = ValAxis.new(@yValAxId, @xValAxId)
      super(frame, options)
      @series_type = ScatterSeries
      @d_lbls = nil
      parse_options options
    end

    # see #scatterStyle
    def scatterStyle=(v)
      Axlsx.validate_scatter_style(v)
      @scatterStyle = v
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      super(str) do |str_inner|
        str_inner << '<c:scatterChart>'
        str_inner << '<c:scatterStyle val="' << scatterStyle.to_s << '"/>'
        str_inner << '<c:varyColors val="' << vary_colors.to_s << '"/>'
        @series.each { |ser| ser.to_xml_string(str_inner) }
        d_lbls.to_xml_string(str) if @d_lbls
        str_inner << '<c:axId val="' << @xValAxId.to_s << '"/>'
        str_inner << '<c:axId val="' << @yValAxId.to_s << '"/>'
        str_inner << '</c:scatterChart>'
        @xValAxis.to_xml_string str_inner
        @yValAxis.to_xml_string str_inner
      end
      str
    end
  end
end
