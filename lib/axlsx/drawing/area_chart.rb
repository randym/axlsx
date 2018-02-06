# encoding: UTF-8
module Axlsx

  # The AreaChart is a two dimentional line chart (who would have guessed?) that you can add to your worksheet.
  # @example Creating a chart
  #   # This example creates a line in a single sheet.
  #   require "rubygems" # if that is your preferred way to manage gems!
  #   require "axlsx"
  #
  #   p = Axlsx::Package.new
  #   ws = p.workbook.add_worksheet
  #   ws.add_row ["This is a chart with no data in the sheet"]
  #
  #   chart = ws.add_chart(Axlsx::AreaChart, :start_at=> [0,1], :end_at=>[0,6], :title=>"Most Popular Pets")
  #   chart.add_series :data => [1, 9, 10], :labels => ["Slimy Reptiles", "Fuzzy Bunnies", "Rottweiler"]
  #
  # @see Worksheet#add_chart
  # @see Worksheet#add_row
  # @see Chart#add_series
  # @see Series
  # @see Package#serialize
  class AreaChart < Chart

    # the category axis
    # @return [CatAxis]
    def cat_axis
      axes[:cat_axis]
    end
    alias :catAxis :cat_axis

    # the category axis
    # @return [ValAxis]
    def val_axis
      axes[:val_axis]
    end
    alias :valAxis :val_axis

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
      super(frame, options)
      @series_type = AreaSeries
      @d_lbls = nil
    end

    # @see grouping
    def grouping=(v)
      RestrictionValidator.validate "AreaChart.grouping", [:percentStacked, :standard, :stacked], v
      @grouping = v
    end

    # The node name to use in serialization. As AreaChart is used as the
    # base class for Liine3DChart we need to be sure to serialize the
    # chart based on the actual class type and not a fixed node name.
    # @return [String]
    def node_name
      path = self.class.to_s
      if i = path.rindex('::')
        path = path[(i+2)..-1]
      end
      path[0] = path[0].chr.downcase
      path
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      super(str) do
        str << ("<c:" << node_name << ">")
        str << ('<c:grouping val="' << grouping.to_s << '"/>')
        str << ('<c:varyColors val="' << vary_colors.to_s << '"/>')
        @series.each { |ser| ser.to_xml_string(str) }
        @d_lbls.to_xml_string(str) if @d_lbls
        yield if block_given?
        axes.to_xml_string(str, :ids => true)
        str << ("</c:" << node_name << ">")
        axes.to_xml_string(str)
      end
    end

    # The axes for this chart. AreaCharts have a category and value
    # axis.
    # @return [Axes]
    def axes
      @axes ||= Axes.new(:cat_axis => CatAxis, :val_axis => ValAxis)
    end
  end
end
