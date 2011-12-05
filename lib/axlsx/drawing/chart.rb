# -*- coding: utf-8 -*-
module Axlsx
  # A Chart is the superclass for specific charts
  # @note Worksheet#add_chart is the recommended way to create charts for your worksheets.
  # @see README for examples
  class Chart


    # The 3D view properties for the chart
    attr_reader :view3D

    # A reference to the graphic frame that owns this chart
    # @return [GraphicFrame]
    attr_reader :graphic_frame

    # A collection of series objects that are applied to the chart
    # @return [SimpleTypedList]
    attr_reader :series

    # The type of series to use for this chart.
    # @return [Series]
    attr_reader :series_type

    #TODO data labels!
    #attr_reader :dLabls

    # The title object for the chart.
    # @return [Title]
    attr_reader :title

    # The style for the chart. 
    # see ECMA Part 1 §21.2.2.196
    # @return [Integer]
    attr_reader :style

    # Show the legend in the chart
    # @return [Boolean]
    attr_reader :show_legend
   
    # Creates a new chart object
    # @param [GraphicalFrame] frame The frame that holds this chart.
    # @option options [Cell, String] title
    # @option options [Boolean] show_legend
    def initialize(frame, options={})
      @style = 2
      @graphic_frame=frame
      @graphic_frame.anchor.drawing.worksheet.workbook.charts << self
      @series = SimpleTypedList.new Series
      @show_legend = true
      @series_type = Series
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
      start_at *options[:start_at] if options[:start_at]
      end_at *options[:end_at] if options[:start_at]
      yield self if block_given?
    end

    # The index of this chart in the workbooks charts collection
    # @return [Integer]
    def index
      @graphic_frame.anchor.drawing.worksheet.workbook.charts.index(self)
    end

    # The part name for this chart
    # @return [String]
    def pn
      "#{CHART_PN % (index+1)}"
    end

    # The title object for the chart.
    # @param [String, Cell] v
    # @return [Title]
    def title=(v) 
      v = Title.new(v) if v.is_a?(String) || v.is_a?(Cell)
      DataTypeValidator.validate "#{self.class}.title", Title, v
      @title = v
    end

    # Show the legend in the chart
    # @param [Boolean] v
    # @return [Boolean]
    def show_legend=(v) Axlsx::validate_boolean(v); @show_legend = v; end


    # The style for the chart. 
    # see ECMA Part 1 §21.2.2.196
    # @param [Integer] v must be between 1 and 48
    def style=(v) DataTypeValidator.validate "Chart.style", Integer, v, lambda { |arg| arg >= 1 && arg <= 48 }; @style = v; end

    # backwards compatibility to allow chart.to and chart.from access to anchor markers
    # @note This will be disconinued in version 2.0.0. Please use the end_at method
    def to
      @graphic_frame.anchor.to
    end

    # backwards compatibility to allow chart.to and chart.from access to anchor markers
    # @note This will be disconinued in version 2.0.0. please use the start_at method
    def from
      @graphic_frame.anchor.from
    end

    # Adds a new series to the chart's series collection.
    # @return [Series]
    # @see Series
    def add_series(options={})
      @series_type.new(self, options)
      @series.last
    end

    # Chart Serialization
    # serializes the chart
    def to_xml
      builder = Nokogiri::XML::Builder.new(:encoding => ENCODING) do |xml|
        xml.send('c:chartSpace',:'xmlns:c' => XML_NS_C, :'xmlns:a' => XML_NS_A) {
          xml.send('c:date1904', :val=>Axlsx::Workbook.date1904)
          xml.send('c:style', :val=>style)
          xml.send('c:chart') {
            @title.to_xml(xml) unless @title.nil?
            @view3D.to_xml(xml) unless @view3D.nil?
            xml.send('c:plotArea') {
              xml.send('c:layout')
              yield xml if block_given?
            }
            if @show_legend
              xml.send('c:legend') {
                xml.send('c:legendPos', :val => "r")
                xml.send('c:layout')
              }
            end
          }
        }
      end
      builder.to_xml
    end

    # This is a short cut method to set the start anchor position
    # If you need finer granularity in positioning use
    # graphic_frame.anchor.from.colOff / rowOff
    # @param [Integer] x The column
    # @param [Integer] y The row
    # @return [Marker]
    def start_at(x, y=0)
      x, y = *parse_coord_args(x, y)
      @graphic_frame.anchor.from.col = x
      @graphic_frame.anchor.from.row = y
    end

    # This is a short cut method to set the end anchor position
    # If you need finer granularity in positioning use
    # graphic_frame.anchor.to.colOff / rowOff
    # @param [Integer] x The column
    # @param [Integer] y The row
    # @return [Marker]
    def end_at(x, y=0)
      x, y = *parse_coord_args(x, y)
      @graphic_frame.anchor.to.col = x
      @graphic_frame.anchor.to.row = y
    end

    private

    def parse_coord_args(x, y=0)
      if x.is_a?(String)
        x, y = *Axlsx::name_to_indices(x)
      end
      if x.is_a?(Cell)
        x, y = *x.pos
      end
      if x.is_a?(Array)
        x, y = *x
      end
      [x, y]
    end

    def view3D=(v) DataTypeValidator.validate "#{self.class}.view3D", View3D, v; @view3D = v; end

  end
end
