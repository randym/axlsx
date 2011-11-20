module Axlsx


  # The Pie3DChart is a three dimentional piechart (who would have guessed?) that you can add to your worksheet.
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
  #   chart = ws.add_chart(Axlsx::Pie3DChart, :start_at=> [0,1], :end_at=>[0,6], :title=>"Most Popular Pets")
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
  class Pie3DChart < Chart


    # Creates a new pie chart object
    # @param [Workbook] workbook The workbook that owns this chart.
    # @option options [Cell, String] title
    def initialize(workbook, options={})
      super(workbook, options)
      # this charts series type
      @series_type = PieSeries
      @view3D = View3D.new(:rotX => 30, :perspective => 30)
    end

    # Serializes the pie chart
    # @return [String]
    def to_xml
      super() do |xml|
        xml.send('c:pie3DChart') {
          xml.send('c:varyColors', :val=>1)
          @series.each { |ser| ser.to_xml(xml) }
        }                      
      end
    end
  end
end
