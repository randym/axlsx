module Axlsx

  class BarChart < Chart

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

    #grouping for a column, line, or area chart.
    # must be one of  [:percentStacked, :clustered, :standard, :stacked]
    # @return [Symbol]
    def grouping
      @grouping ||= :clustered
    end

    # space between bar or column clusters, as a percentage of the bar or column width.
    # @return [String]
    def gap_width
      @gap_width ||= 150
    end
    alias :gapWidth :gap_width

    # validation regex for gap amount percent
    GAP_AMOUNT_PERCENT = /0*(([0-9])|([1-9][0-9])|([1-4][0-9][0-9])|500)%/

    def initialize(frame, options={})
      @gap_width = nil
      super(frame, options)
      @series_type = BarSeries
      @d_lbls = nil
      @vary_colors = true
    end

    def to_xml_string(str = '')
      super(str) do |str_inner|
        str_inner << '<c:barChart>'
        str_inner << '<c:barDir val="' << bar_dir.to_s << '"/>'
        str_inner << '<c:grouping val="' << grouping.to_s << '"/>'
        str_inner << '<c:varyColors val="' << vary_colors.to_s << '"/>'
        @series.each { |ser| ser.to_xml_string(str_inner) }
        @d_lbls.to_xml_string(str_inner) if @d_lbls
        str_inner << '<c:gapWidth val="' << @gap_width.to_s << '"/>' unless @gap_width.nil?
        axes.to_xml_string(str_inner, :ids => true)
        str_inner << '</c:barChart>'
        axes.to_xml_string(str_inner)
      end
    end

    # The direction of the bars in the chart
    # must be one of [:bar, :col]
    def bar_dir=(v)
      RestrictionValidator.validate "BarChart.bar_dir", [:bar, :col], v
      @bar_dir = v
    end
    alias :barDir= :bar_dir=

    #grouping for a column, line, or area chart.
    # must be one of  [:percentStacked, :clustered, :standard, :stacked]
    def grouping=(v)
      RestrictionValidator.validate "BarChart.grouping", [:percentStacked, :clustered, :standard, :stacked], v
      @grouping = v
    end

    # space between bar or column clusters, as a percentage of the bar or column width.
    def gap_width=(v)
      RegexValidator.validate "BarChart.gap_width", GAP_AMOUNT_PERCENT, v
      @gap_width=(v)
    end
    alias :gapWidth= :gap_width=

    # A hash of axes used by this chart. Bar charts have a value and
    # category axes specified via axes[:val_axes] and axes[:cat_axis]
    # @return [Axes]
    def axes
      @axes ||= Axes.new(:cat_axis => CatAxis, :val_axis => ValAxis)
    end
  end
end