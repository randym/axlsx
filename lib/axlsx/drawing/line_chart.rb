# encoding: UTF-8
module Axlsx

  # The LineChart is a two dimentional line chart (who would have guessed?) that you can add to your worksheet.
  # @example Creating a chart
  #   # This example creates a line in a single sheet.
  #   require "rubygems" # if that is your preferred way to manage gems!
  #   require "axlsx"
  #
  #   p = Axlsx::Package.new
  #   ws = p.workbook.add_worksheet
  #   ws.add_row ["This is a chart with no data in the sheet"]
  #
  #   chart = ws.add_chart(Axlsx::LineChart, :start_at=> [0,1], :end_at=>[0,6], :title=>"Most Popular Pets")
  #   chart.add_series :data => [1, 9, 10], :labels => ["Slimy Reptiles", "Fuzzy Bunnies", "Rottweiler"]
  #
  # @see Worksheet#add_chart
  # @see Worksheet#add_row
  # @see Chart#add_series
  # @see Series
  # @see Package#serialize
  class LineChart < Chart

    # the category axis
    # @return [CatAxis]
    attr_reader :catAxis

    # the category axis
    # @return [ValAxis]
    attr_reader :valAxis

    # the category axis
    # @return [Axis]
    attr_reader :serAxis

    # must be one of  [:percentStacked, :clustered, :standard, :stacked]
    # @return [Symbol]
    attr_reader :grouping

    # Creates a new line chart object
    # @param [GraphicFrame] frame The workbook that owns this chart.
    # @option options [Cell, String] title
    # @option options [Boolean] show_legend
    # @option options [Symbol] grouping
    # @see Chart
    def initialize(frame, options={})
      @vary_colors = false
      @grouping = :standard
      @catAxId = rand(8 ** 8)
      @valAxId = rand(8 ** 8)
      @serAxId = rand(8 ** 8)
      @catAxis = CatAxis.new(@catAxId, @valAxId)
      @valAxis = ValAxis.new(@valAxId, @catAxId)
      @serAxis = SerAxis.new(@serAxId, @valAxId)
      super(frame, options)
      @series_type = LineSeries
      @d_lbls = nil
    end

    # @see grouping
    def grouping=(v)
      RestrictionValidator.validate "Bar3DChart.grouping", [:percentStacked, :standard, :stacked], v
      @grouping = v
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      super(str) do |str_inner|
        str_inner << "<c:#{self.class.name.camelcase(:lower)}>"
        str_inner << '<c:grouping val="' << grouping.to_s << '"/>'
        str_inner << '<c:varyColors val="' << vary_colors.to_s << '"/>'
        @series.each { |ser| ser.to_xml_string(str_inner) }
        @d_lbls.to_xml_string(str) if @d_lbls
        yield str_inner if block_given?
        str_inner << '<c:axId val="' << @catAxId.to_s << '"/>'
        str_inner << '<c:axId val="' << @valAxId.to_s << '"/>'
        str_inner << '<c:axId val="' << @serAxId.to_s << '"/>'
        str_inner << "</c:#{self.class.name.camelcase(:lower)}>"
        @catAxis.to_xml_string str_inner
        @valAxis.to_xml_string str_inner
        @serAxis.to_xml_string str_inner
      end
    end
  end
end
