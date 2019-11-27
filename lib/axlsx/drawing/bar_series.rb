# encoding: UTF-8
# frozen_string_literal: true
module Axlsx
  # A BarSeries defines the title, data and labels for bar charts
  # @note The recommended way to manage series is to use Chart#add_series
  # @see Worksheet#add_chart
  # @see Chart#add_series
  class BarSeries < Series


    # The data for this series.
    # @return [NumDataSource]
    attr_reader :data

    # The labels for this series.
    # @return [Array, SimpleTypedList]
    attr_reader :labels

    # The shabe of the bars or columns
    # must be one of  [:percentStacked, :clustered, :standard, :stacked]
    # @return [Symbol]
    attr_reader :shape

    # An array of rgb colors to apply to your bar chart.
    attr_reader :colors

    # Creates a new series
    # @option options [Array, SimpleTypedList] data
    # @option options [Array, SimpleTypedList] labels
    # @option options [String] title
    # @option options [String] shape
    # @option options [String] colors an array of colors to use when rendering each data point
    # @param [Chart] chart
    def initialize(chart, options={})
      @shape = :box
      @colors = []
      super(chart, options)
      self.labels = AxDataSource.new({:data => options[:labels]}) unless options[:labels].nil?
      self.data = NumDataSource.new(options) unless options[:data].nil?
    end

    # @see colors
    def colors=(v) DataTypeValidator.validate "BarSeries.colors", [Array], v; @colors = v end

    # The shabe of the bars or columns
    # must be one of  [:cone, :coneToMax, :box, :cylinder, :pyramid, :pyramidToMax]
    def shape=(v)
      RestrictionValidator.validate "BarSeries.shape", [:cone, :coneToMax, :box, :cylinder, :pyramid, :pyramidToMax], v
      @shape = v
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = String.new)
      super(str) do

        colors.each_with_index do |c, index|
          str << '<c:dPt>'\
                 "<c:idx val=\"#{index}\"/>"\
                 '<c:spPr><a:solidFill>'\
                 "<a:srgbClr val=\"#{c}\"/>"\
                 '</a:solidFill></c:spPr></c:dPt>'
        end

        @labels.to_xml_string(str) unless @labels.nil?
        @data.to_xml_string(str) unless @data.nil?
        # this is actually only required for shapes other than box
        str << "<c:shape val=\"#{shape}\"></c:shape>"
      end
    end

    private

    # assigns the data for this series
    def data=(v) DataTypeValidator.validate "Series.data", [NumDataSource], v; @data = v; end

    # assigns the labels for this series
    def labels=(v) DataTypeValidator.validate "Series.labels", [AxDataSource], v; @labels = v; end

  end

end
