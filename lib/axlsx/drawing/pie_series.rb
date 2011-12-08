module Axlsx
  # A PieSeries defines the data and labels and explosion for pie charts series.
  # @note The recommended way to manage series is to use Chart#add_series
  # @see Worksheet#add_chart
  # @see Chart#add_series
  class PieSeries < Series

    # The data for this series. 
    # @return [SimpleTypedList]
    attr_reader :data

    # The labels for this series.
    # @return [SimpleTypedList]
    attr_reader :labels

    # The explosion for this series
    # @return [Integert]
    attr_reader :explosion

    # Creates a new series
    # @option options [Array, SimpleTypedList] data
    # @option options [Array, SimpleTypedList] labels
    # @option options [String] title
    # @option options [Integer] explosion
    # @param [Chart] chart
    def initialize(chart, options={})
      @explosion = nil
      super(chart, options)
      self.labels = CatAxisData.new(options[:labels]) unless options[:labels].nil?
      self.data = ValAxisData.new(options[:data]) unless options[:data].nil?
    end 
    
    # @see explosion
    def explosion=(v) Axlsx::validate_unsigned_int(v); @explosion = v; end

    # Serializes the series
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      super(xml) do  |xml_inner|
        xml_inner.explosion :val=>@explosion unless @explosion.nil?
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
