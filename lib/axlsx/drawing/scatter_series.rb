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

    # Serializes the series
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      super(xml) do |xml_inner|
        @xData.to_xml(xml_inner) unless @xData.nil?
        @yData.to_xml(xml_inner) unless @yData.nil?
      end
    end

  end
end
