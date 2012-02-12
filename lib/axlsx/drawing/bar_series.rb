# encoding: UTF-8
module Axlsx
  # A BarSeries defines the title, data and labels for bar charts
  # @note The recommended way to manage series is to use Chart#add_series
  # @see Worksheet#add_chart
  # @see Chart#add_series
  class BarSeries < Series

  
    # The data for this series. 
    # @return [Array, SimpleTypedList]
    attr_reader :data

    # The labels for this series.
    # @return [Array, SimpleTypedList]
    attr_reader :labels

    # The shabe of the bars or columns
    # must be one of  [:percentStacked, :clustered, :standard, :stacked]
    # @return [Symbol]
    attr_reader :shape

    # Creates a new series
    # @option options [Array, SimpleTypedList] data
    # @option options [Array, SimpleTypedList] labels
    # @option options [String] title
    # @option options [String] shape
    # @param [Chart] chart
    def initialize(chart, options={})
      @shape = :box
      super(chart, options)
      self.labels = CatAxisData.new(options[:labels]) unless options[:labels].nil?
      self.data = ValAxisData.new(options[:data]) unless options[:data].nil?
    end 

    # The shabe of the bars or columns
    # must be one of  [:percentStacked, :clustered, :standard, :stacked]
    def shape=(v) 
      RestrictionValidator.validate "BarSeries.shape", [:cone, :coneToMax, :box, :cylinder, :pyramid, :pyramidToMax], v
      @shape = v
    end

    # Serializes the series
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      super(xml) do |xml_inner|
        @labels.to_xml(xml_inner) unless @labels.nil?
        @data.to_xml(xml_inner) unless @data.nil?
        xml_inner.shape :val=>@shape
      end      
    end

    
    private 

    # assigns the data for this series
    def data=(v) DataTypeValidator.validate "Series.data", [SimpleTypedList], v; @data = v; end

    # assigns the labels for this series
    def labels=(v) DataTypeValidator.validate "Series.labels", [SimpleTypedList], v; @labels = v; end

  end

end
