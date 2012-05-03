# encoding: UTF-8
module Axlsx

  # A ScatterSeries defines the x and y position of data in the chart
  # @note The recommended way to manage series is to use Chart#add_series
  # @see Worksheet#add_chart
  # @see Chart#add_series
  # @see examples/example.rb
  class ScatterSeries < Series

    # The x data for this series.
    # @return [NamedAxisData]
    attr_reader :xData

    # The y data for this series.
    # @return [NamedAxisData]
    attr_reader :yData

    # The fill color for this series.
    # Red, green, and blue is expressed as sequence of hex digits, RRGGBB. A perceptual gamma of 2.2 is used.
    # @return [String]
    attr_reader :color

    # Creates a new ScatterSeries
    def initialize(chart, options={})
      @xData, @yData = nil
      super(chart, options)
      @xData = AxDataSource.new(:tag_name => :xVal, :data => options[:xData]) unless options[:xData].nil?
      @yData = NumDataSource.new({:tag_name => :yVal, :data => options[:yData]}) unless options[:yData].nil?
    end

    # @see color
    def color=(v)
      @color = v
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      super(str) do |inner_str|
        # needs to override the super color here to push in ln/and something else!
        if color
          str << '<c:spPr><a:solidFill>'
          str << '<a:srgbClr val="' << color << '"/>'
          str << '</a:solidFill>'
          str << '<a:ln><a:solidFill>'
          str << '<a:srgbClr val="' << color << '"/></a:solidFill></a:ln>'
          str << '</c:spPr>'
          str << '<c:marker>'
          str << '<c:spPr><a:solidFill>'
          str << '<a:srgbClr val="' << color << '"/>'
          str << '</a:solidFill>'
          str << '<a:ln><a:solidFill>'
          str << '<a:srgbClr val="' << color << '"/></a:solidFill></a:ln>'
          str << '</c:spPr>'
          str << '</c:marker>'
        end
        @xData.to_xml_string(inner_str) unless @xData.nil?
        @yData.to_xml_string(inner_str) unless @yData.nil?
      end
      str
    end
  end
end
