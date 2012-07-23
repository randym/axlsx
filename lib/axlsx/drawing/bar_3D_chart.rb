# encoding: UTF-8
module Axlsx

  # The Bar3DChart is a three dimentional barchart (who would have guessed?) that you can add to your worksheet.
  # @see Worksheet#add_chart
  # @see Chart#add_series
  # @see Package#serialize
  # @see README for an example
  class Bar3DChart < Chart

    # the category axis
    # @return [CatAxis]
    attr_reader :cat_axis
    alias :catAxis :cat_axis

    # the value axis
    # @return [ValAxis]
    attr_reader :val_axis
    alias :valAxis :val_axis

    # The direction of the bars in the chart
    # must be one of [:bar, :col]
    # @return [Symbol]
    attr_reader :bar_dir
    alias :barDir :bar_dir

    # space between bar or column clusters, as a percentage of the bar or column width.
    # @return [String]
    attr_reader :gap_depth
    alias :gapDepth :gap_depth

    # space between bar or column clusters, as a percentage of the bar or column width.
    # @return [String]
    attr_reader :gap_width 
    alias :gapWidth :gap_width

    #grouping for a column, line, or area chart.
    # must be one of  [:percentStacked, :clustered, :standard, :stacked]
    # @return [Symbol]
    attr_reader :grouping

    # The shabe of the bars or columns
    # must be one of  [:cone, :coneToMax, :box, :cylinder, :pyramid, :pyramidToMax]
    # @return [Symbol]
    attr_reader :shape

    # validation regex for gap amount percent
    GAP_AMOUNT_PERCENT = /0*(([0-9])|([1-9][0-9])|([1-4][0-9][0-9])|500)%/

    # Creates a new bar chart object
    # @param [GraphicFrame] frame The workbook that owns this chart.
    # @option options [Cell, String] title
    # @option options [Boolean] show_legend
    # @option options [Symbol] bar_dir
    # @option options [Symbol] grouping
    # @option options [String] gap_width
    # @option options [String] gap_depth
    # @option options [Symbol] shape
    # @option options [Integer] rot_x
    # @option options [String] h_percent
    # @option options [Integer] rot_y
    # @option options [String] depth_percent
    # @option options [Boolean] r_ang_ax
    # @option options [Integer] perspective
    # @see Chart
    # @see View3D
    def initialize(frame, options={})
      @bar_dir = :bar
      @grouping = :clustered
      @shape = :box
      @gap_width = 150
      @gap_width, @gap_depth, @shape = nil, nil, nil
      @cat_ax_id = rand(8 ** 8)
      @val_ax_id = rand(8 ** 8)
      @cat_axis = CatAxis.new(@cat_ax_id, @val_ax_id)
      @val_axis = ValAxis.new(@val_ax_id, @cat_ax_id, :tick_lbl_pos => :low, :ax_pos => :l)
      super(frame, options)
      @series_type = BarSeries
      @view_3D = View3D.new({:r_ang_ax=>1}.merge(options))
      @d_lbls = nil
    end

    # The direction of the bars in the chart
    # must be one of [:bar, :col]
    def bar_dir=(v)
      RestrictionValidator.validate "Bar3DChart.bar_dir", [:bar, :col], v
      @bar_dir = v
    end
    alias :barDir= :bar_dir=

    #grouping for a column, line, or area chart.
    # must be one of  [:percentStacked, :clustered, :standard, :stacked]
    def grouping=(v)
      RestrictionValidator.validate "Bar3DChart.grouping", [:percentStacked, :clustered, :standard, :stacked], v
      @grouping = v
    end

    # space between bar or column clusters, as a percentage of the bar or column width.
    def gap_width=(v)
      RegexValidator.validate "Bar3DChart.gap_width", GAP_AMOUNT_PERCENT, v
      @gap_width=(v)
    end
    alias :gapWidth= :gap_width=

    # space between bar or column clusters, as a percentage of the bar or column width.
    def gap_depth=(v)
      RegexValidator.validate "Bar3DChart.gap_didth", GAP_AMOUNT_PERCENT, v
      @gap_depth=(v)
    end
    alias :gapDepth= :gap_depth=

    # The shabe of the bars or columns
    # must be one of  [:cone, :coneToMax, :box, :cylinder, :pyramid, :pyramidToMax]
    def shape=(v)
      RestrictionValidator.validate "Bar3DChart.shape", [:cone, :coneToMax, :box, :cylinder, :pyramid, :pyramidToMax], v
      @shape = v
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      super(str) do |str_inner|
        str_inner << '<c:bar3DChart>'
        str_inner << '<c:barDir val="' << bar_dir.to_s << '"/>'
        str_inner << '<c:grouping val="' << grouping.to_s << '"/>'
        str_inner << '<c:varyColors val="1"/>'
        @series.each { |ser| ser.to_xml_string(str_inner) }
        @d_lbls.to_xml_string(str) if @d_lbls
        str_inner << '<c:gapWidth val="' << @gap_width.to_s << '"/>' unless @gap_width.nil?
        str_inner << '<c:gapDepth val="' << @gap_depth.to_s << '"/>' unless @gap_depth.nil?
        str_inner << '<c:shape val="' << @shape.to_s << '"/>' unless @shape.nil?
        str_inner << '<c:axId val="' << @cat_ax_id.to_s << '"/>'
        str_inner << '<c:axId val="' << @val_ax_id.to_s << '"/>'
        str_inner << '<c:axId val="0"/>'
        str_inner << '</c:bar3DChart>'
        @cat_axis.to_xml_string str_inner
        @val_axis.to_xml_string str_inner
      end
    end
  end
end
