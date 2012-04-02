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

    # Creates a new ScatterSeries
    def initialize(chart, options={})
      @xData, @yData = nil
      super(chart, options)

      @xData = NamedAxisData.new("xVal", options[:xData]) unless options[:xData].nil?
      @yData = NamedAxisData.new("yVal", options[:yData]) unless options[:yData].nil?
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      super(str) do |inner_str|
        @xData.to_xml_string(inner_str) unless @xData.nil?
        @yData.to_xml_string(inner_str) unless @yData.nil?
      end
      str
    end
  end
end
