# encoding: UTF-8
module Axlsx
  # A Series defines the common series attributes and is the super class for all concrete series types.
  # @note The recommended way to manage series is to use Chart#add_series
  # @see Worksheet#add_chart
  # @see Chart#add_series
  class Series

    include Axlsx::OptionsParser

    # The chart that owns this series
    # @return [Chart]
    attr_reader :chart

    # The title of the series
    # @return [SeriesTitle]
    attr_reader :title

    # Creates a new series
    # @param [Chart] chart
    # @option options [Integer] order
    # @option options [String] title
    def initialize(chart, options={})
      @order = nil
      self.chart = chart
      @chart.series << self
      parse_options options
    end

    # The index of this series in the chart's series.
    # @return [Integer]
    def index
      @chart.series.index(self)
    end

    # The order of this series in the chart's series. By default the order is the index of the series.
    # @return [Integer]
    def order
      @order || index
    end

    # @see order
    def order=(v)  Axlsx::validate_unsigned_int(v); @order = v; end

    # @see title
    def title=(v)
      v = SeriesTitle.new(v) if v.is_a?(String) || v.is_a?(Cell)
      DataTypeValidator.validate "#{self.class}.title", SeriesTitle, v
      @title = v
    end

    private

    # assigns the chart for this series
    def chart=(v)  DataTypeValidator.validate "Series.chart", Chart, v; @chart = v; end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<c:ser>'
      str << ('<c:idx val="' << index.to_s << '"/>')
      str << ('<c:order val="' << (order || index).to_s << '"/>')
      title.to_xml_string(str) unless title.nil?
      yield if block_given?
      str << '</c:ser>'
    end
  end
end
