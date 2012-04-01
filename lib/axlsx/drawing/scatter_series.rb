# encoding: UTF-8
module Axlsx
  class ScatterSeries < Series
    # The x data for this series.
    # @return [NamedAxisData]
    attr_reader :xData

    # The y data for this series.
    # @return [NamedAxisData]
    attr_reader :yData

    def initialize(chart, options={})
      @xData, @yData = nil
      super(chart, options)

      @xData = NamedAxisData.new("xVal", options[:xData]) unless options[:xData].nil?
      @yData = NamedAxisData.new("yVal", options[:yData]) unless options[:yData].nil?
    end

    def to_xml_string(str = '')
      super(str) do |inner_str|
        @xData.to_xml_string(inner_str) unless @xData.nil?
        @yData.to_xml_string(inner_str) unless @yData.nil?
      end
      str
    end
  end
end
