module Axlsx
  require 'axlsx/drawing/series.rb'
  require 'axlsx/drawing/pie_series.rb'
  require 'axlsx/drawing/bar_series.rb'  

  require 'axlsx/drawing/axis.rb'
  require 'axlsx/drawing/cat_axis.rb'
  require 'axlsx/drawing/val_axis.rb'
  require 'axlsx/drawing/view_3D.rb'
  require 'axlsx/drawing/scaling.rb'
  require 'axlsx/drawing/title.rb'

  require 'axlsx/drawing/graphic_frame.rb'
  require 'axlsx/drawing/marker.rb'
  require 'axlsx/drawing/two_cell_anchor.rb'

  require 'axlsx/drawing/chart.rb'
  require 'axlsx/drawing/pie_3D_chart.rb'
  require 'axlsx/drawing/bar_3D_chart.rb'

  # A Drawing is a canvas for charts. Each worksheet has a single drawing that can specify multiple anchors which reference charts.
  # @note The recommended way to manage drawings is to use the Worksheet.add_chart method, specifying the chart class, start and end marker locations. 
  # @see Worksheet#add_chart
  # @see TwoCellAnchor
  # @see Chart
  class Drawing

    # The worksheet that owns the drawing
    # @return [Worksheet]
    attr_reader :worksheet

   
    # A collection of anchors for this drawing
    # @return [SimpleTypedList]
    attr_reader :anchors

    # An array of charts that are associated with this drawing's anchors
    # @return [Array]
    attr_reader :charts

    # The index of this drawing in the owning workbooks's drawings collection.
    # @return [Integer]
    attr_reader :index

    # The relation reference id for this drawing
    # @return [String]
    attr_reader :rId

    # The part name for this drawing
    # @return [String]
    attr_reader :pn

    # The relational part name for this drawing
    # @return [String]
    attr_reader :rels_pn

    # The drawing's relationships.
    # @return [Relationships]
    attr_reader :relationships

    # Creates a new Drawing object
    # @param [Worksheet] worksheet The worksheet that owns this drawing
    def initialize(worksheet)
      DataTypeValidator.validate "Drawing.worksheet", Worksheet, worksheet
      @worksheet = worksheet
      @worksheet.workbook.drawings << self
      @anchors = SimpleTypedList.new TwoCellAnchor
    end
    

    # Adds a chart to the drawing. 
    # @note The recommended way to manage charts is to use Worksheet.add_chart.
    # @param [Chart] chart_type The class of the chart to be added to the drawing
    # @param [Hash] options
    def add_chart(chart_type, options={})
      DataTypeValidator.validate "Drawing.chart_type", Chart, chart_type 
      TwoCellAnchor.new(self, chart_type, options)
      @anchors.last.graphic_frame.chart
    end
    
    def charts
      @anchors.map { |a| a.graphic_frame.chart }
    end

    def index
      @worksheet.workbook.drawings.index(self)
    end

    def rId
      "rId#{index+1}"
    end

    def pn
      "#{DRAWING_PN % (index+1)}"
    end
    
    def rels_pn
      "#{DRAWING_RELS_PN % (index+1)}"
    end

    def relationships
      r = Relationships.new
      @anchors.each do |anchor|
        chart = anchor.graphic_frame.chart
        r << Relationship.new(CHART_R, "../#{chart.pn}")
      end
      r
    end

    # Serializes the drawing
    # @return [String]
    def to_xml
      builder = Nokogiri::XML::Builder.new(:encoding => ENCODING) do |xml|
        xml.send('xdr:wsDr', :'xmlns:xdr'=>XML_NS_XDR, :'xmlns:a'=>XML_NS_A) {
          anchors.each {|anchor| anchor.to_xml(xml) }
        }        
      end
      builder.to_xml
    end
  end
end
