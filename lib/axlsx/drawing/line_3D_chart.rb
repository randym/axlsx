# encoding: UTF-8
module Axlsx

  # The Line3DChart is a three dimentional line chart (who would have guessed?) that you can add to your worksheet.
  # @example Creating a chart
  #   # This example creates a line in a single sheet.
  #   require "rubygems" # if that is your preferred way to manage gems!
  #   require "axlsx"
  #
  #   p = Axlsx::Package.new
  #   ws = p.workbook.add_worksheet
  #   ws.add_row ["This is a chart with no data in the sheet"]
  #
  #   chart = ws.add_chart(Axlsx::Line3DChart, :start_at=> [0,1], :end_at=>[0,6], :t#itle=>"Most Popular Pets")
  #   chart.add_series :data => [1, 9, 10], :labels => ["Slimy Reptiles", "Fuzzy Bunnies", "Rottweiler"]
  #
  # @see Worksheet#add_chart
  # @see Worksheet#add_row
  # @see Chart#add_series
  # @see Series
  # @see Package#serialize
  class Line3DChart < Axlsx::LineChart

    # space between bar or column clusters, as a percentage of the bar or column width.
    # @return [String]
    attr_reader :gap_depth
    alias :gapDepth :gap_depth

    # validation regex for gap amount percent
    GAP_AMOUNT_PERCENT = /0*(([0-9])|([1-9][0-9])|([1-4][0-9][0-9])|500)%/

    # the category axis
    # @return [Axis]
    def ser_axis
      axes[:ser_axis]
    end
    alias :serAxis :ser_axis

    # Creates a new line chart object
    # @option options [String] gap_depth
    # @see Chart
    # @see lineChart
    # @see View3D
    def initialize(frame, options={})
      @gap_depth = nil
      @view_3D = View3D.new({:r_ang_ax=>1}.merge(options))
      super(frame, options)
      axes.add_axis :ser_axis, SerAxis
    end


    # @see gapDepth
    def gap_depth=(v)
      RegexValidator.validate "Line3DChart.gapWidth", GAP_AMOUNT_PERCENT, v
      @gap_depth=(v)
    end
    alias :gapDepth= :gap_depth=

      # Serializes the object
      # @param [String] str
      # @return [String]
      def to_xml_string(str = '')
        super(str) do
          str << ('<c:gapDepth val="' << @gap_depth.to_s << '"/>') unless @gap_depth.nil?
        end
      end
  end
end
