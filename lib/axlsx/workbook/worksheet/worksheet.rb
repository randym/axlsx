# -*- coding: utf-8 -*-
module Axlsx
  
  # The Worksheet class represents a worksheet in the workbook. 
  class Worksheet
    
    # The name of the worksheet
    # @return [String]
    attr_accessor :name

    # The workbook that owns this worksheet
    # @return [Workbook]
    attr_reader :workbook

    # The worksheet relationships. This is managed automatically by the worksheet
    # @return [Relationships]
    attr_reader :relationships

    # The rows in this worksheet
    # @note The recommended way to manage rows is Worksheet#add_row
    # @return [SimpleTypedList]
    # @see Worksheet#add_row
    attr_reader :rows
    
    # The drawing associated with this worksheet.
    # @note the recommended way to work with drawings and charts is Worksheet#add_chart
    # @return [Drawing]
    # @see Worksheet#add_chart
    attr_reader :drawing

    # An array of content based calculated column widths.
    # @note a single auto fit data item is a hash with :longest => [String] and :sz=> [Integer] members.
    # @return [Array] of Hash 
    attr_reader :auto_fit_data

    # The part name of this worksheet
    # @return [String]
    attr_reader :pn

    # The relationship part name of this worksheet
    # @return [String]
    attr_reader :rels_pn

    # The relationship Id of thiw worksheet
    # @return [String]
    attr_reader :rId

    # The index of this worksheet in the owning Workbook's worksheets list.
    # @return [Integer]
    attr_reader :index

    # TODO Merge Cells
    # attr_reader :merge_cells
    
    # Creates a new worksheet.
    # @note the recommended way to manage worksheets is Workbook#add_worksheet
    # @see Workbook#add_worksheet
    # @option options [String] name The name of this sheet.
    def initialize(wb, options={})
      @rows = SimpleTypedList.new Row
      self.workbook = wb
      @workbook.worksheets << self
      @auto_fit_data = []
      self.name = options[:name] || "Sheet" + (index+1).to_s
      @magick_draw = Magick::Draw.new
    end

    
    def name=(v) DataTypeValidator.validate "Worksheet.name", String, v; @name=v end
    
    def pn
      "#{WORKSHEET_PN % (index+1)}"
    end

    def rels_pn
      "#{WORKSHEET_RELS_PN % (index+1)}"
    end

    def rId
      "rId#{index+1}"
    end
    
    def index
      @workbook.worksheets.index(self)
    end

    def drawing
      @drawing || @drawing = Axlsx::Drawing.new(self)
    end

    # Adds a row to the worksheet and updates auto fit data
    # @return [Row]
    # @option options [Array] values
    # @option options [Array, Symbol] types 
    # @option options [Array, Integer] style 
    def add_row(values=[], options={})
      Row.new(self, values, options)
      update_auto_fit_data @rows.last.cells
      yield @rows.last if block_given?
      @rows.last
    end

    # Adds a chart to this worksheets drawing. This is the recommended way to create charts for your worksheet. This method wraps the complexity of dealing with ooxml drawing, anchors, markers graphic frames chart objects and all the other dirty details. 
    # @param [Class] chart_type
    # @option options [Array] start_at
    # @option options [Array] end_at
    # @option options [Cell, String] title
    # @option options [Boolean] show_legend
    # @option options [Integer] style 
    # @note each chart type also specifies additional options 
    # @see Chart
    # @see Pie3DChart
    # @see Bar3DChart
    # @see Line3DChart
    # @see README for examples
    def add_chart(chart_type, options={})
      chart = drawing.add_chart(chart_type, options)
      yield chart if block_given?
      chart
    end

    # Serializes the worksheet document
    # @return [String]
    def to_xml
      builder = Nokogiri::XML::Builder.new(:encoding => ENCODING) do |xml|
        xml.worksheet(:xmlns => XML_NS, :'xmlns:r' => XML_NS_R) {
          if @auto_fit_data.size > 0
            xml.cols {
              @auto_fit_data.each_with_index do |col, index|
                min_max = index+1
                xml.col(:min=>min_max, :max=>min_max, :width => auto_width(col), :customWidth=>"true")
              end
            }
          end
          xml.sheetData {
            @rows.each do |row|
              row.to_xml(xml)
            end
          }
          xml.drawing :"r:id"=>"rId1" if @drawing          
        }
      end
      builder.to_xml(:indent=>0, :save_with=>0)
    end

    # The worksheet's relationships.
    def relationships
        r = Relationships.new
        r << Relationship.new(DRAWING_R, "../#{@drawing.pn}") if @drawing
        r
    end

    private 

    # assigns the owner workbook for this worksheet
    def workbook=(v) DataTypeValidator.validate "Worksheet.workbook", Workbook, v; @workbook = v; end

    # Updates auto fit data. 
    # Autofit data attempts to determine the cell in a column that has the greatest width by comparing the length of the text multiplied by the size of the font.
    # @return [Array] of Cell objects
    # @param [Array] cells an array of cells
    def update_auto_fit_data(cells)
      styles = self.workbook.styles
      cellXfs, fonts = styles.cellXfs, styles.fonts
      sz = fonts[0].sz

      cells.each_with_index do |item, index|
        col = @auto_fit_data[index] || {:longest=>"", :sz=>sz} 
        cell_xf = cellXfs[item.style]
        font = fonts[cell_xf.fontId || 0]
        sz = font.sz || sz

        if (col[:longest].scan(/./mu).size * col[:sz]) < (item.value.to_s.scan(/./mu).size * sz)
          col[:sz] =  sz
          col[:longest] = item.value.to_s
        end
        @auto_fit_data[index] = col
      end
      cells
    end

    # Determines the proper width for a column based on content.
    # @note 
    #  From ECMA docs
    #   Column width measured as the number of characters of the maximum digit width of the numbers 0 .. 9 as 
    #   rendered in the normal style's font. There are 4 pixels of margin padding (two on each side), plus 1 pixel padding for the gridlines.
    #   width = Truncate([{Number of Characters} * {Maximum Digit Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256
    # @return [Float]
    # @param [Hash] A hash of auto_fit_data 
    def auto_width(col)
      mdw = 6.0 # maximum digit with is always 6.0 in testable fonts so instead of beating RMagick every time, I am hardcoding it here.
      mdw_count = 0 
      best_guess = 1.5  #direct testing shows the results of the documented formula to be a bit too small. This is a best guess scaling
      font_scale = col[:sz].to_f / (self.workbook.styles.fonts[0].sz.to_f || 11.0)
      col[:longest].scan(/./mu).each do |i|
        mdw_count +=1 if @magick_draw.get_type_metrics(i).width >= mdw 
      end
      ((mdw_count * mdw + 5) / mdw * 256) / 256.0 * best_guess * font_scale      
    end   
  end
end
