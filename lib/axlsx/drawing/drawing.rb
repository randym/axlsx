# encoding: UTF-8
module Axlsx
  require 'axlsx/drawing/title.rb'
  require 'axlsx/drawing/series_title.rb'
  require 'axlsx/drawing/series.rb'
  require 'axlsx/drawing/pie_series.rb'
  require 'axlsx/drawing/bar_series.rb'  
  require 'axlsx/drawing/line_series.rb'  
  require 'axlsx/drawing/scatter_series.rb'

  require 'axlsx/drawing/scaling.rb'
  require 'axlsx/drawing/axis.rb'
  require 'axlsx/drawing/ser_axis.rb'
  require 'axlsx/drawing/cat_axis.rb'
  require 'axlsx/drawing/val_axis.rb'

  require 'axlsx/drawing/cat_axis_data.rb'
  require 'axlsx/drawing/val_axis_data.rb'
  require 'axlsx/drawing/named_axis_data.rb'

  require 'axlsx/drawing/marker.rb'
 
  require 'axlsx/drawing/one_cell_anchor.rb'
  require 'axlsx/drawing/two_cell_anchor.rb'
  require 'axlsx/drawing/graphic_frame.rb'

  require 'axlsx/drawing/view_3D.rb'
  require 'axlsx/drawing/chart.rb'
  require 'axlsx/drawing/pie_3D_chart.rb'
  require 'axlsx/drawing/bar_3D_chart.rb'
  require 'axlsx/drawing/line_3D_chart.rb'
  require 'axlsx/drawing/scatter_chart.rb'

  require 'axlsx/drawing/picture_locking.rb'
  require 'axlsx/drawing/pic.rb'
  require 'axlsx/drawing/hyperlink.rb'

  # A Drawing is a canvas for charts. Each worksheet has a single drawing that manages anchors.
  # The anchors reference the charts via graphical frames. This is not a trivial relationship so please do follow the advice in the note.
  # @note The recommended way to manage drawings is to use the Worksheet.add_chart method.
  # @see Worksheet#add_chart
  # @see Chart
  # see README for an example of how to create a chart.
  class Drawing

    # The worksheet that owns the drawing
    # @return [Worksheet]
    attr_reader :worksheet
   
    # A collection of anchors for this drawing
    # only TwoCellAnchors are supported in this version
    # @return [SimpleTypedList]
    attr_reader :anchors

    # Creates a new Drawing object
    # @param [Worksheet] worksheet The worksheet that owns this drawing
    def initialize(worksheet)
      DataTypeValidator.validate "Drawing.worksheet", Worksheet, worksheet
      @worksheet = worksheet
      @worksheet.workbook.drawings << self
      @anchors = SimpleTypedList.new [TwoCellAnchor, OneCellAnchor]
    end

    # Adds an image to the chart
    # @note The recommended way to manage images is to use Worksheet.add_image. Please refer to that method for documentation.
    # @see Worksheet#add_image
    def add_image(options={})
      OneCellAnchor.new(self, options)
      @anchors.last.object
    end

    # Adds a chart to the drawing.
    # @note The recommended way to manage charts is to use Worksheet.add_chart. Please refer to that method for documentation.
    # @see Worksheet#add_chart
    def add_chart(chart_type, options={})
      TwoCellAnchor.new(self, options)
      @anchors.last.add_chart(chart_type, options)
    end

    # An array of charts that are associated with this drawing's anchors
    # @return [Array]
    def charts
      charts = @anchors.select { |a| a.object.is_a?(GraphicFrame) }
      charts.map { |a| a.object.chart }
    end

    # An array of hyperlink objects associated with this drawings images
    # @return [Array]
    def hyperlinks
      links = self.images.select { |a| a.hyperlink.is_a?(Hyperlink) }
      links.map { |a| a.hyperlink }
    end

    # An array of image objects that are associated with this drawing's anchors
    # @return [Array]
    def images
      images = @anchors.select { |a| a.object.is_a?(Pic) }
      images.map { |a| a.object }
    end

    # The index of this drawing in the owning workbooks's drawings collection.
    # @return [Integer]
    def index
      @worksheet.workbook.drawings.index(self)
    end

    # The relation reference id for this drawing
    # @return [String]
    def rId
      "rId#{index+1}"
    end

    # The part name for this drawing
    # @return [String]
    def pn
      "#{DRAWING_PN % (index+1)}"
    end

    # The relational part name for this drawing
    # @return [String]
    def rels_pn
      "#{DRAWING_RELS_PN % (index+1)}"
    end

    # The drawing's relationships.
    # @return [Relationships]
    def relationships
      r = Relationships.new
      charts.each do |chart|
        r << Relationship.new(CHART_R, "../#{chart.pn}")
      end
      images.each do |image|
        r << Relationship.new(IMAGE_R, "../#{image.pn}")
      end
      hyperlinks.each do |hyperlink|
        r << Relationship.new(HYPERLINK_R, hyperlink.href, :target_mode => :External)
      end
      r
    end

    # Serializes the drawing
    # @return [String]
    def to_xml
      builder = Nokogiri::XML::Builder.new(:encoding => ENCODING) do |xml|
        xml.send('xdr:wsDr', :'xmlns:xdr'=>XML_NS_XDR, :'xmlns:a'=>XML_NS_A, :'xmlns:c'=>XML_NS_C) {
          anchors.each {|anchor| anchor.to_xml(xml) }
        }        
      end
      builder.to_xml(:save_with => 0)
    end
  end
end
