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
    def cat_axis
      axes[:cat_axis]
    end
    alias :catAxis :cat_axis

    # the values axis
    # @return [ValAxis]
    def val_axis
      axes[:val_axis]
    end
    alias :valAxis :val_axis

    # the secondary category axis
    # @return [sec_cat_axis]
    def sec_cat_axis
      axes[:sec_cat_axis]
    end
    alias :secCatAxis :sec_cat_axis

    # the secondary values axis
    # @return [sec_val_axis]
    def sec_val_axis
      axes[:sec_val_axis]
    end
    alias :secValAxis :sec_val_axis

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
      @series_type = LineSeries
      @d_lbls = nil
    end

    # @see grouping
    def grouping=(v)
      RestrictionValidator.validate "LineChart.grouping", [:percentStacked, :standard, :stacked], v
      @grouping = v
    end

    # The node name to use in serialization. As LineChart is used as the
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
      if @series.all? {|s| s.on_primary_axis} then
        # Only a primary val axis
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
      else
        # Two value axes
        super(str) do
          # First axis
          str << ("<c:" << node_name << ">")
          str << ('<c:grouping val="' << grouping.to_s << '"/>')
          str << ('<c:varyColors val="' << vary_colors.to_s << '"/>')
          @series.select {|s| s.on_primary_axis}.each { |s| s.to_xml_string(str) }
          @d_lbls.to_xml_string(str) if @d_lbls
          yield if block_given?
          str << ('<c:axId val="' << axes[:cat_axis].id.to_s << '"/>')
          str << ('<c:axId val="' << axes[:val_axis].id.to_s << '"/>')
          str << ("</c:" << node_name << ">")

          # Secondary axis
          str << ("<c:" << node_name << ">")
          str << ('<c:grouping val="' << grouping.to_s << '"/>')
          str << ('<c:varyColors val="' << vary_colors.to_s << '"/>')
          @series.select {|s| !s.on_primary_axis}.each { |s| s.to_xml_string(str) }
          @d_lbls.to_xml_string(str) if @d_lbls
          yield if block_given?
          str << ('<c:axId val="' << axes[:sec_cat_axis].id.to_s << '"/>')
          str << ('<c:axId val="' << axes[:sec_val_axis].id.to_s << '"/>')
          str << ("</c:" << node_name << ">")

          # The axes
          axes.to_xml_string(str)
        end
      end
    end

    # The axes for this chart. LineCharts have a category and value
    # axis. If any series is on the secondary axis we will have two
    # category and two value axes.
    # @return [Axes]
    def axes
      if @axes.nil? then
        # add the normal axes
        @axes = Axes.new(:cat_axis => CatAxis, :val_axis => ValAxis)

        # add the secondary axes if needed
        if @series.any? {|s| !s.on_primary_axis} then
          if @axes[:sec_cat_axis].nil? then
            @axes.add_axis(:sec_cat_axis, Axlsx::CatAxis)
            sec_cat_axis = @axes[:sec_cat_axis]
            sec_cat_axis.ax_pos = :b
            sec_cat_axis.delete = 1
            sec_cat_axis.gridlines = false
          end
          if @axes[:sec_val_axis].nil? then
            @axes.add_axis(:sec_val_axis, Axlsx::ValAxis)
            sec_val_axis = @axes[:sec_val_axis]
            sec_val_axis.ax_pos = :r
            sec_val_axis.gridlines = false
            sec_val_axis.crosses = :max
          end
        end
      end

      # return
      @axes
    end
  end
end
