# encoding: UTF-8
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

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      super(str) do |str_inner|
        str_inner << '<c:explosion val="' << @explosion << '"/>' unless @explosion.nil?
        @labels.to_xml_string str_inner unless @labels.nil?
        @data.to_xml_string str_inner unless @data.nil?
      end
      str
    end

    private

    # assigns the data for this series
    def data=(v) DataTypeValidator.validate "Series.data", [SimpleTypedList], v; @data = v; end

    # assigns the labels for this series
    def labels=(v) DataTypeValidator.validate "Series.labels", [SimpleTypedList], v; @labels = v; end

  end

end
