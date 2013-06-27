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
    def cat_axis
      axes[:cat_axis]
    end
    alias :catAxis :cat_axis

    # the value axis
    # @return [ValAxis]
    def val_axis
      axes[:val_axis]
    end
    alias :valAxis :val_axis

    # The direction of the bars in the chart
    # must be one of [:bar, :col]
    # @return [Symbol]
    def bar_dir
      @bar_dir ||= :bar
    end
    alias :barDir :bar_dir

    # space between bar or column clusters, as a percentage of the bar or column width.
    # @return [String]
    attr_reader :gap_depth
    alias :gapDepth :gap_depth

    # space between bar or column clusters, as a percentage of the bar or column width.
    # @return [String]
    def gap_width
      @gap_width ||= 150
    end
    alias :gapWidth :gap_width
    
    #grouping for a column, line, or area chart.
    # must be one of  [:percentStacked, :clustered, :standard, :stacked]
    # @return [Symbol]
    def grouping
      @grouping ||= :clustered
    end

    # The shabe of the bars or columns
    # must be one of  [:cone, :coneToMax, :box, :cylinder, :pyramid, :pyramidToMax]
    # @return [Symbol]
    def shape
      @shape ||= :box
    end

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
      @vary_colors = true
      @gap_width, @gap_depth, @shape = nil, nil, nil
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
        str_inner << '<c:varyColors val="' << vary_colors.to_s << '"/>'
        @series.each { |ser| ser.to_xml_string(str_inner) }
        @d_lbls.to_xml_string(str_inner) if @d_lbls
        str_inner << '<c:gapWidth val="' << @gap_width.to_s << '"/>' unless @gap_width.nil?
        str_inner << '<c:gapDepth val="' << @gap_depth.to_s << '"/>' unless @gap_depth.nil?
        str_inner << '<c:shape val="' << @shape.to_s << '"/>' unless @shape.nil?
        axes.to_xml_string(str_inner, :ids => true)
        str_inner << '</c:bar3DChart>'
        axes.to_xml_string(str_inner)
      end
    end

    # A hash of axes used by this chart. Bar charts have a value and
    # category axes specified via axes[:val_axes] and axes[:cat_axis]
    # @return [Axes]
    def axes
     @axes ||= Axes.new(:cat_axis => CatAxis, :val_axis => ValAxis)
    end
  end
end
