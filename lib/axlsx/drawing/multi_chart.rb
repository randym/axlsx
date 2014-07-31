# encoding: UTF-8
module Axlsx

  # The MultiChart is a wrapper for multiple charts that you can add to your worksheet.
  # @see Worksheet#add_chart
  # @see Chart#add_series
  # @see Package#serialize
  # @see README for an example
  class MultiChart < Chart

    module OverrideToXmlString
      def to_xml_string(str = '')
        yield if block_given?
        str
      end
    end

    # the charts this multi chart contains
    # @return [SimpleTypedList]
    def sub_charts
      @sub_charts
    end

    # Creates a new multi chart object
    # @param [GraphicFrame] frame The workbook that owns this chart.
    # @see Chart
    def initialize(frame, options = {})
      super(frame, options)
      @sub_charts ||= SimpleTypedList.new(Chart)
    end

    # Add a sub_chart
    # @see Worksheet.add_chart
    def add_sub_chart(chart_type, options = {})
      chart = chart_type.new(@graphic_frame, options)
      @graphic_frame.anchor.drawing.worksheet.workbook.charts.pop
      yield chart if block_given?
      @sub_charts << chart
      chart
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      base_index = 0
      super(str) do
        sub_charts.each do |sub_chart|
          # Yes, I went there
          sub_chart.instance_eval{ class << self; self; end }.superclass.send(:include, OverrideToXmlString)
          # Yes, I really went there
          sub_chart.series.each do |ser|
            ser.define_singleton_method(:index) do
              base_index
            end
            base_index += 1
          end
          sub_chart.to_xml_string(str)
        end
      end
      str
    end
  end
end
