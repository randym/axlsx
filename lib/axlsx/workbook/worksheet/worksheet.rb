# -*- coding: utf-8 -*-
module Axlsx
  
  # The Worksheet class represents a worksheet in the workbook. 
  class Worksheet
    
    # The name of the worksheet
    # @return [String]
    attr_reader :name

    # The workbook that owns this worksheet
    # @return [Workbook]
    attr_reader :workbook


    # The rows in this worksheet
    # @note The recommended way to manage rows is Worksheet#add_row
    # @return [SimpleTypedList]
    # @see Worksheet#add_row
    attr_reader :rows

    # An array of content based calculated column widths.
    # @note a single auto fit data item is a hash with :longest => [String] and :sz=> [Integer] members.
    # @return [Array] of Hash 
    attr_reader :auto_fit_data

    # TODO Merge Cells
    # attr_reader :merge_cells
    
    # Creates a new worksheet.
    # @note the recommended way to manage worksheets is Workbook#add_worksheet
    # @see Workbook#add_worksheet
    # @option options [String] name The name of this sheet.
    def initialize(wb, options={})
      @drawing = nil
      @rows = SimpleTypedList.new Row
      self.workbook = wb
      @workbook.worksheets << self
      @auto_fit_data = []
      self.name = options[:name] || "Sheet" + (index+1).to_s
      @magick_draw = Magick::Draw.new
      @cols = SimpleTypedList.new Cell
    end

    # Returns the cell or cells defined using excel style A1:B3 references.
    # @param [String] cell_def the string defining the cell or range of cells
    # @return [Cell, Array]
    def [](cell_def)
      parts = cell_def.split(':')
      first = name_to_cell parts[0]

      if parts.size == 1
        first
      else
        cells = []
        last = name_to_cell(parts[1])
        rows[(first.row.index..last.row.index)].each do |r|
          r.cells[(first.index..last.index)].each do |c|
            cells << c
          end
        end
        cells
      end
    end

    # returns the column and row index for a named based cell
    # @param [String] name The cell or cell range to return. "A1" will return the first cell of the first row.
    # @return [Cell]
    def name_to_cell(name)
      col_index, row_index = *Axlsx::name_to_indices(name)
      r = rows[row_index]
      r.cells[col_index] if r
    end

    # The name of the worksheet
    # @param [String] v
    def name=(v) 
      DataTypeValidator.validate "Worksheet.name", String, v
      sheet_names = @workbook.worksheets.map { |s| s.name }
      raise ArgumentError, (ERR_DUPLICATE_SHEET_NAME % v) if sheet_names.include?(v) 
      @name=v 
    end

    # The part name of this worksheet
    # @return [String]
    def pn
      "#{WORKSHEET_PN % (index+1)}"
    end

    # The relationship part name of this worksheet
    # @return [String]
    def rels_pn
      "#{WORKSHEET_RELS_PN % (index+1)}"
    end

    # The relationship Id of thiw worksheet
    # @return [String]
    def rId
      "rId#{index+1}"
    end

    # The index of this worksheet in the owning Workbook's worksheets list.
    # @return [Integer]
    def index
      @workbook.worksheets.index(self)
    end

    # The drawing associated with this worksheet.
    # @note the recommended way to work with drawings and charts is Worksheet#add_chart
    # @return [Drawing]
    # @see Worksheet#add_chart
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

    # Set the style for cells in a specific row
    # @param [Integer] index or range of indexes in the table
    # @param [Integer] the cellXfs index
    # @option options [Integer] col_offset only cells after this column will be updated.
    # @note You can also specify the style in the add_row call
    # @see Worksheet#add_row
    # @see README.md for an example
    def row_style(index, style, options={})
      offset = options.delete(:col_offset) || 0
      rs = @rows[index]
      if rs.is_a?(Array)
        rs.each { |r| r.cells[(offset..-1)].each { |c| c.style = style } }
      else
        rs.cells[(offset..-1)].each { |c| c.style = style }
      end
    end

    # returns the sheet data as columnw
    def cols
      @rows.transpose
    end


    # Set the style for cells in a specific column
    # @param [Integer] index the index of the column
    # @param [Integer] the cellXfs index
    # @option options [Integer] row_offset only cells after this column will be updated.
    # @note You can also specify the style for specific columns in the call to add_row by using an array for the :styles option
    # @see Worksheet#add_row
    # @see README.md for an example
    def col_style(index, style, options={})
      offset = options.delete(:row_offset) || 0
      @rows[(offset..-1)].each do |r| 
        cells = r.cells[index]
        if cells.is_a?(Array)
          cells.each { |c| c.style = style }
        else
          cells.style = style
        end
      end
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

    # Adds a media item to the worksheets drawing
    # @param [Class] media_type
    # @option options [] unknown
    def add_image(options={})
      image = drawing.add_image(options)
      yield image if block_given?
      image
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
      builder.to_xml
    end

    # The worksheet relationships. This is managed automatically by the worksheet
    # @return [Relationships]
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
        next if item.value.is_a?(String) && item.value.start_with?('=')
        col = @auto_fit_data[index] || {:longest=>"", :sz=>sz} 
        cell_xf = cellXfs[item.style]
        font = fonts[cell_xf.fontId || 0]
        sz = item.sz || font.sz || fonts[0].sz
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
    #   width = Truncate([!{Number of Characters} * !{Maximum Digit Width} + !{5 pixel padding}]/{Maximum Digit Width}*256)/256
    # @return [Float]
    # @param [Hash] A hash of auto_fit_data 
    def auto_width(col)
      mdw = 6.0 # maximum digit with is always 6.0 with RMagick's default font
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
