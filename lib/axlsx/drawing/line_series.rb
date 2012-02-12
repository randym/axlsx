# encoding: UTF-8
module Axlsx
  # A LineSeries defines the title, data and labels for line charts
  # @note The recommended way to manage series is to use Chart#add_series
  # @see Worksheet#add_chart
  # @see Chart#add_series
  class LineSeries < Series
  
    # The data for this series. 
    # @return [ValAxisData]
    attr_reader :data

    # The labels for this series.
    # @return [CatAxisData]
    attr_reader :labels

    # Creates a new series
    # @option options [Array, SimpleTypedList] data
    # @option options [Array, SimpleTypedList] labels
    # @param [Chart] chart
    def initialize(chart, options={})
      @labels, @data = nil, nil
      super(chart, options)
      @labels = CatAxisData.new(options[:labels]) unless options[:labels].nil?
      @data = ValAxisData.new(options[:data]) unless options[:data].nil?
    end 

    # Serializes the series
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      super(xml) do |xml_inner|
        @labels.to_xml(xml_inner) unless @labels.nil?
        @data.to_xml(xml_inner) unless @data.nil?
      end      
    end

    private 

    # assigns the data for this series
    def data=(v) DataTypeValidator.validate "Series.data", [SimpleTypedList], v; @data = v; end

    # assigns the labels for this series
    def labels=(v) DataTypeValidator.validate "Series.labels", [SimpleTypedList], v; @labels = v; end

  end
end
