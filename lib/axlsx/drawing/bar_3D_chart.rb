module Axlsx

  # The Bar3DChart is a three dimentional barchart (who would have guessed?) that you can add to your worksheet.
  # @example Creating a chart
  #   # This example creates two charts in a single sheet.
  #   # The first uses data directly fed to the sheet, while the second references cells withing the worksheet for data.
  #
  #   require "rubygems" # if that is your preferred way to manage gems!
  #   require "axlsx"
  #
  #   p = Axlsx::Package.new
  #   ws = p.workbook.add_worksheet
  #   ws.add_row :values => ["This is a chart with no data in the sheet"]
  #
  #   chart = ws.add_chart(Axlsx::Bar3DChart, :start_at=> [0,1], :end_at=>[0,6], :title=>"Most Popular Pets")
  #   chart.add_series :data => [1, 9, 10], :labels => ["Slimy Reptiles", "Fuzzy Bunnies", "Rottweiler"]
  # 
  #   ws.add_row :values => ["This chart uses the data below"]
  #   title_row = ws.add_row :values => ["Least Popular Pets"]
  #   label_row = ws.add_row :values => ["", "Dry Skinned Reptiles", "Bald Cats", "Violent Parrots"]
  #   data_row = ws.add_row :values => ["Votes", 6, 4, 1]
  #   
  #   chart = ws.add_chart(Axlsx::Pie3DChart, :start_at => [0,11], :end_at =>[0,16], :title => title_row.cells.last)
  #   chart.add_series :data => data_row.cells[(1..-1)], :labels => label_row.cells  
  #
  #   f = File.open('example_pie_3d_chart.xlsx', 'w')
  #   p.serialize(f)
  #
  # @see Worksheet#add_chart
  # @see Worksheet#add_row
  # @see Chart#add_series
  # @see Series
  # @see Package#serialize
  class Bar3DChart < Chart

    # the category axis
    # @return [CatAxis]
    attr_reader :catAxis

    # the category axis
    # @return [ValAxis]
    attr_reader :valAxis

    # The direction of the bars in the chart
    # must be one of [:bar, :col]
    # @return [Symbol]
    attr_accessor :barDir

    # space between bar or column clusters, as a percentage of the bar or column width.
    # @return [String]
    attr_accessor :gapDepth

    # space between bar or column clusters, as a percentage of the bar or column width.
    # @return [String]
    attr_accessor :gapWidth

    #grouping for a column, line, or area chart.
    # must be one of  [:percentStacked, :clustered, :standard, :stacked]
    # @return [Symbol]
    attr_accessor :grouping

    # The shabe of the bars or columns
    # must be one of  [:percentStacked, :clustered, :standard, :stacked]
    # @return [Symbol]
    attr_accessor :shape

    # validation regex for gap amount percent
    GAP_AMOUNT_PERCENT = /0*(([0-9])|([1-9][0-9])|([1-4][0-9][0-9])|500)%/
    
    # Creates a new bar chart object
    # @param [GraphicFrame] frame The workbook that owns this chart.
    # @option options [Cell, String] title
    # @option options [Boolean] show_legend
    # @option options [Symbol] barDir
    # @option options [Symbol] grouping
    # @option options [String] gapWidth
    # @option options [String] gapDepth
    # @option options [Symbol] shape
    def initialize(frame, options={})
      super(frame, options)      
      @series_type = BarSeries
      @barDir = :bar
      @grouping = :clustered
      @catAxId = rand(8 ** 8)
      @valAxId = rand(8 ** 8)
      @catAxis = CatAxis.new(@catAxId, @valAxId)
      @valAxis = ValAxis.new(@valAxId, @catAxId)
      @view3D = View3D.new(:rAngAx=>1)
    end
    
    def barDir=(v) 
      RestrictionValidator.validate "Bar3DChart.barDir", [:bar, :col], v
      @barDir = v
    end


    def grouping=(v)
      RestrictionValidator.validate "Bar3DChart.grouping", [:percentStacked, :clustered, :standard, :stacked], v
      @grouping = v
    end

    def gapWidth=(v)
      RegexValidator.validate "Bar3DChart.gapWidth", GAP_AMOUNT_PERCENT, v
      @gapWidth=(v)
    end

    def gapDepth=(v)
      RegexValidator.validate "Bar3DChart.gapWidth", GAP_AMOUNT_PERCENT, v
      @gapDepth=(v)
    end

    def shape=(v) 
      RestrictionValidator.validate "Bar3DChart.shape", [:cone, :coneToMax, :box, :cylinder, :pyramid, :pyramidToMax], v
      @shape = v
    end
    
    # Serializes the bar chart
    # @return [String]
    def to_xml
      super() do |xml|
        xml.send('c:bar3DChart') {
          xml.send('c:barDir', :val => barDir)
          xml.send('c:grouping', :val=>grouping)
          xml.send('c:varyColors', :val=>1)
          @series.each { |ser| ser.to_xml(xml) }
          xml.send('c:gapWidth', :val=>@gapWidth) unless @gapWidth.nil?
          xml.send('c:gapDepth', :val=>@gapDepth) unless @gapDepth.nil?
          xml.send('c:shape', :val=>@shape) unless @shape.nil?
          xml.send('c:axId', :val=>@catAxId)
          xml.send('c:axId', :val=>@valAxId)
          xml.send('c:axId', :val=>0)
        }
        @catAxis.to_xml(xml)
        @valAxis.to_xml(xml)        
      end
    end    
  end  
end
