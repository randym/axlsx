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

    # An array of rgb colors to apply to your bar chart.
    attr_reader :colors

    # Creates a new series
    # @option options [Array, SimpleTypedList] data
    # @option options [Array, SimpleTypedList] labels
    # @option options [String] title
    # @option options [Integer] explosion
    # @param [Chart] chart
    def initialize(chart, options={})
      @explosion = nil
      @colors = []
      super(chart, options)
      self.labels = AxDataSource.new(:data => options[:labels]) unless options[:labels].nil?
      self.data = NumDataSource.new(options) unless options[:data].nil?
    end

    # @see colors
    def colors=(v) DataTypeValidator.validate "BarSeries.colors", [Array], v; @colors = v end

    # @see explosion
    def explosion=(v) Axlsx::validate_unsigned_int(v); @explosion = v; end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      super(str) do |str_inner|
        str_inner << '<c:explosion val="' << @explosion << '"/>' unless @explosion.nil?
        colors.each_with_index do |c, index|
          str << '<c:dPt>'
          str << '<c:idx val="' << index.to_s << '"/>'
          str << '<c:spPr><a:solidFill>'
          str << '<a:srgbClr val="' << c << '"/>'
          str << '</a:solidFill></c:spPr></c:dPt>'
        end
        @labels.to_xml_string str_inner unless @labels.nil?
        @data.to_xml_string str_inner unless @data.nil?
      end
      str
    end

    private

    # assigns the data for this series
    def data=(v) DataTypeValidator.validate "Series.data", [NumDataSource], v; @data = v; end

    # assigns the labels for this series
    def labels=(v) DataTypeValidator.validate "Series.labels", [AxDataSource], v; @labels = v; end

  end

end
