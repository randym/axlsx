# encoding: UTF-8
module Axlsx

  # The Worksheet class represents a worksheet in the workbook.
  class Worksheet

    # The name of the worksheet
    # @return [String]
    attr_reader :name

    # The workbook that owns this worksheet
    # @return [Workbook]
    attr_reader :workbook

    # The tables in this worksheet
    # @return [Array] of Table
    attr_reader :tables

    # The rows in this worksheet
    # @note The recommended way to manage rows is Worksheet#add_row
    # @return [SimpleTypedList]
    # @see Worksheet#add_row
    attr_reader :rows

    # An array of content based calculated column widths.
    # @note a single auto fit data item is a hash with :longest => [String] and :sz=> [Integer] members.
    # @return [Array] of Hash
    attr_reader :auto_fit_data

    # An array of merged cell ranges e.d "A1:B3"
    # Content and formatting is read from the first cell.
    # @return Array
    attr_reader :merged_cells

    # An range that excel will apply an autfilter to "A1:B3"
    # This will turn filtering on for the cells in the range.
    # The first row is considered the header, while subsequent rows are considerd to be data.
    # @return Array
    attr_reader :auto_filter

    # Indicates if the worksheet should show gridlines or not
    # @return Boolean
    attr_reader :show_gridlines


    # Indicates if the worksheet is selected in the workbook
    # It is possible to have more than one worksheet selected, however it might cause issues
    # in some older versions of excel when using copy and paste.
    # @return Boolean
    attr_reader :selected

    # Indicates if the worksheet should print in a single page
    # @return Boolean
    attr_reader :fit_to_page

    # Page margins for printing the worksheet.
    # @example
    #      wb = Axlsx::Package.new.workbook
    #      # using options when creating the worksheet.
    #      ws = wb.add_worksheet :page_margins => {:left => 1.9, :header => 0.1}
    #
    #      # use the set method of the page_margins object
    #      ws.page_margins.set(:bottom => 3, :footer => 0.7)
    #
    #      # set page margins in a block
    #      ws.page_margins do |margins|
    #        margins.right = 6
    #        margins.top = 0.2
    #      end
    # @see PageMargins#initialize
    # @return [PageMargins]
    def page_margins
      @page_margins ||= PageMargins.new
      yield @page_margins if block_given?
      @page_margins

    end

    # Creates a new worksheet.
    # @note the recommended way to manage worksheets is Workbook#add_worksheet
    # @see Workbook#add_worksheet
    # @option options [String] name The name of this worksheet.
    # @option options [Hash] page_margins A hash containing page margins for this worksheet. @see PageMargins
    # @option options [Boolean] show_gridlines indicates if gridlines should be shown for this sheet.
    def initialize(wb, options={})
      self.workbook = wb
      @workbook.worksheets << self

      @drawing = @page_margins = @auto_filter = nil
      @merged_cells = []
      @auto_fit_data = []

      @selected = false
      @show_gridlines = true
      self.name = "Sheet" + (index+1).to_s
      @page_margins = PageMargins.new options[:page_margins] if options[:page_margins]

      @rows = SimpleTypedList.new Row
      @cols = SimpleTypedList.new Cell
      @tables = SimpleTypedList.new Table

      if self.workbook.use_autowidth
        require 'RMagick' unless defined?(Magick)
        @magick_draw = Magick::Draw.new
      else
        @magick_draw = nil
      end

      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end

    end

    # convinience method to access all cells in this worksheet
    # @return [Array] cells
    def cells
      rows.flatten
    end

    # Creates merge information for this worksheet.
    # Cells can be merged by calling the merge_cells method on a worksheet.
    # @example This would merge the three cells C1..E1    #
    #        worksheet.merge_cells "C1:E1"
    #        # you can also provide an array of cells to be merged
    #        worksheet.merge_cells worksheet.rows.first.cells[(2..4)]
    #        #alternatively you can do it from a single cell
    #        worksheet["C1"].merge worksheet["E1"]
    # @param [Array, string]
    def merge_cells(cells)
      @merged_cells << if cells.is_a?(String)
                         cells
                       elsif cells.is_a?(Array)
                         cells = cells.sort { |x, y| x.r <=> y.r }
                         "#{cells.first.r}:#{cells.last.r}"
                       end
    end


    # The demensions of a worksheet. This is not actually a required element by the spec,
    # but at least a few other document readers expect this for conversion
    # @return [String] the A1:B2 style reference for the first and last row column intersection in the workbook
    def dimension
      "#{rows.first.cells.first.r}:#{rows.last.cells.last.r}"
    end

    # Indicates if gridlines should be shown in the sheet.
    # This is true by default.
    # @return [Boolean]
    def show_gridlines=(v)
      Axlsx::validate_boolean v
      @show_gridlines = v
    end

    # @see selected
    # @return [Boolean]
    def selected=(v)
      Axlsx::validate_boolean v
      @selected = v
    end


    # Indicates if gridlines should be shown in the sheet.
    # This is true by default.
    # @return [Boolean]
    def fit_to_page=(v)
      Axlsx::validate_boolean v
      @fit_to_page = v
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
    # The name of a worksheet must be unique in the workbook, and must not exceed 31 characters
    # @param [String] v
    def name=(v)
      DataTypeValidator.validate "Worksheet.name", String, v
      raise ArgumentError, (ERR_SHEET_NAME_TOO_LONG % v) if v.size > 31
      sheet_names = @workbook.worksheets.map { |s| s.name }
      raise ArgumentError, (ERR_DUPLICATE_SHEET_NAME % v) if sheet_names.include?(v)
      @name=v
    end

    # The auto filter range for the worksheet
    # @param [String] v
    # @see auto_filter
    def auto_filter=(v)
      DataTypeValidator.validate "Worksheet.auto_filter", String, v
      @auto_filter = v
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

    # Adds a row to the worksheet and updates auto fit data.
    # @example - put a vanilla row in your spreadsheet
    #     ws.add_row [1, 'fish on my pl', '8']
    #
    # @example - specify a fixed width for a column in your spreadsheet
    #     # The first column will ignore the content of this cell when calculating column autowidth.
    #     # The second column will include this text in calculating the columns autowidth
    #     # The third cell will set a fixed with of 80 for the column.
    #     # If you need to un-fix a column width, use :auto. That will recalculate the column width based on all content in the column
    #
    #     ws.add_row ['I wish', 'for a fish', 'on my fish wish dish'], :widths=>[:ignore, :auto, 80]
    #
    # @example - specify a fixed height for a row
    #     ws.add_row ['I wish', 'for a fish', 'on my fish wish dish'], :height => 40
    #
    # @example - create and use a style for all cells in the row
    #     blue = ws.styles.add_style :color => "#00FF00"
    #     ws.add_row [1, 2, 3], :style=>blue
    #
    # @example - only style some cells
    #     blue = ws.styles.add_style :color => "#00FF00"
    #     red = ws.styles.add_style :color => "#FF0000"
    #     big = ws.styles.add_style :sz => 40
    #     ws.add_row ["red fish", "blue fish", "one fish", "two fish"], :style=>[red, blue, nil, big] # the last nil is optional
    #
    #
    # @example - force the second cell to be a float value
    #     ws.add_row [3, 4, 5], :types => [nil, :float]
    #
    # @example - use << alias
    #     ws << [3, 4, 5], :types => [nil, :float]
    #
    # @see Worksheet#column_widths
    # @return [Row]
    # @option options [Array] values
    # @option options [Array, Symbol] types
    # @option options [Array, Integer] style
    # @option options [Array] widths each member of the widths array will affect how auto_fit behavies.
    # @option options [Float] height the row's height (in points)
    def add_row(values=[], options={})
      Row.new(self, values, options)
      update_auto_fit_data @rows.last.cells, options.delete(:widths) || []
      yield @rows.last if block_given?
      @rows.last
    end

    alias :<< :add_row

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
        next unless cells
        if cells.is_a?(Array)
          cells.each { |c| c.style = style }
        else
          cells.style = style
        end
      end
    end

    # This is a helper method that Lets you specify a fixed width for multiple columns in a worksheet in one go.
    # Axlsx is sparse, so if you have not set data for a column, you cannot set the width.
    # Setting a fixed column width to nil will revert the behaviour back to calculating the width for you.
    # @example This would set the first and third column widhts but leave the second column in autofit state.
    #      ws.column_widths 7.2, nil, 3
    # @note For updating only a single column it is probably easier to just set ws.auto_fit_data[col_index][:fixed] directly
    # @param [Integer|Float|Fixnum|nil] values
    def column_widths(*args)
      args.each_with_index do |value, index|
        raise ArgumentError, "Invalid column specification" unless index < @auto_fit_data.size
        Axlsx::validate_unsigned_numeric(value) unless value == nil
        @auto_fit_data[index][:fixed] = value
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

    # needs documentation
    def add_table(ref, options={})
      table = Table.new(ref, self, options)
      @tables << table
      yield table if block_given?
      table
    end

    # Adds a media item to the worksheets drawing
    # @param [Class] media_type
    # @option options [] unknown
    def add_image(options={})
      image = drawing.add_image(options)
      yield image if block_given?
      image
    end

    def to_xml_string
      str = "<worksheet xmlns=\"%s\" xmlns:r=\"%s\">" % [XML_NS, XML_NS_R]
      str.concat "<sheetPr><pageSetUpPr fitToPage=\"%s\"></pageSetUpPr></sheetPr>" % fit_to_page if fit_to_page
      str.concat "<dimension ref=\"%s\"></dimension>" % dimension unless rows.size == 0
      str.concat "<sheetViews><sheetView tabSelected='%s' workbookViewId='0' showGridLines='%s'><selection activeCell=\"A1\" sqref=\"A1\"/></sheetView></sheetViews>" % [@selected, show_gridlines]

      if @auto_fit_data.size > 0
        str.concat "<cols>"
        @auto_fit_data.each_with_index do |col, index|
          min_max = index+1
          str.concat "<col min='%s' max='%s' width='%s' customWidth='1'></col>" % [min_max, min_max, auto_width(col)]
        end
        str.concat '</cols>'
      end

      str.concat '<sheetData>'
      @rows.each_with_index { |row, index| row.to_xml_string(index, str) }
      str.concat '</sheetData>'
      str.concat page_margins.to_xml_string if @page_margins
      str.concat "<autoFilter ref='%s'></autoFilter>" % @auto_filter if @auto_filter
      str.concat "<mergeCells count='%s'>%s</mergeCells>" % [@merged_cells.size, @merged_cells.reduce('') { |memo, obj| "<mergeCell ref='%s'></mergeCell>" % obj } ] unless @merged_cells.empty?
      str.concat "<drawing r:id='rId1'></drawing>" if @drawing
      unless @tables.empty?
        str.concat "<tableParts count='%s'>%s</tableParts>" % [@tables.size, @tables.reduce('') { |memo, obj| memo += "<tablePart r:id='%s'/>" % obj.rId }]
      end
      str + '</worksheet>'
    end

    # Serializes the worksheet document
    # @return [String]
    def to_xml
      builder = Nokogiri::XML::Builder.new(:encoding => ENCODING) do |xml|
        xml.worksheet(:xmlns => XML_NS,
                      :'xmlns:r' => XML_NS_R) {
          xml.sheetPr {
            xml.pageSetUpPr :fitToPage => fit_to_page if fit_to_page
          }
          # another patch for the folks at rubyXL as thier parser depends on this optional element.
          xml.dimension :ref=>dimension unless rows.size == 0
          # this is required by rubyXL, spec says who cares - but it seems they didnt notice
          # grouping issue resolved by keeping tabSelected set to 0
          xml.sheetViews {
            xml.sheetView(:tabSelected => @selected, :workbookViewId => 0, :showGridLines => show_gridlines) {
              xml.selection :activeCell=>"A1", :sqref => "A1"
            }
          }

          if @auto_fit_data.size > 0
            xml.cols {
              @auto_fit_data.each_with_index do |col, index|
                min_max = index+1
                xml.col(:min=>min_max, :max=>min_max, :width => auto_width(col), :customWidth=>1)
              end
            }
          end
          xml.sheetData {
            @rows.each do |row|
              row.to_xml(xml)
            end
          }
          xml.autoFilter :ref=>@auto_filter if @auto_filter
          xml.mergeCells(:count=>@merged_cells.size) { @merged_cells.each { | mc | xml.mergeCell(:ref=>mc) } } unless @merged_cells.empty?
          page_margins.to_xml(xml) if @page_margins
          xml.drawing :"r:id"=>"rId1" if @drawing
          unless @tables.empty?
            xml.tableParts(:count => @tables.length) {
              @tables.each do |table|
                xml.tablePart :'r:id' => table.rId
              end
            }
          end
        }
      end
      builder.to_xml(:save_with => 0)
    end

    # The worksheet relationships. This is managed automatically by the worksheet
    # @return [Relationships]
    def relationships
        r = Relationships.new
        @tables.each do |table|
          r << Relationship.new(TABLE_R, "../#{table.pn}")
        end
        r << Relationship.new(DRAWING_R, "../#{@drawing.pn}") if @drawing
        r
    end

    # Returns the cell or cells defined using excel style A1:B3 references.
    # @param [String|Integer] cell_def the string defining the cell or range of cells, or the rownumber
    # @return [Cell, Array]

    def [] (cell_def)
      return rows[cell_def] if cell_def.is_a?(Integer)

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

    private

    # assigns the owner workbook for this worksheet
    def workbook=(v) DataTypeValidator.validate "Worksheet.workbook", Workbook, v; @workbook = v; end

    # Updates auto fit data.
    # We store an auto_fit_data item for each column. when a row is added we multiple the font size by the length of the text to
    # attempt to identify the longest cell in the column. This is not 100% accurate as it needs to take into account
    # any formatting that will be applied to the data, as well as the actual rendering size when the length and size is equal
    # for two cells.

    # @return [Array] of Cell objects
    # @param [Array] cells an array of cells
    # @param [Array] widths an array of cell widths @see Worksheet#add_row
    def update_auto_fit_data(cells, widths=[])
      # TODO delay this until rendering. too much work when we dont know what they are going to do to the sheet.
      styles = self.workbook.styles
      cellXfs, fonts = styles.cellXfs, styles.fonts
      sz = 11
      cells.each_with_index do |item, index|
        col = @auto_fit_data[index] ||= {:longest=>"", :sz=>sz, :fixed=>nil}
        width = widths[index]
        # set fixed width and skip if numeric width is given
        col[:fixed] = width if [Integer, Float, Fixnum].include?(width.class)
        # ignore default column widths and formula
        next if width == :ignore || (item.value.is_a?(String) && item.value.start_with?('='))
        # make sure we can turn that fixed with off!
        col[:fixed] = nil if width == :auto
        next unless self.workbook.use_autowidth

        cell_xf = cellXfs[item.style]
        font = fonts[cell_xf.fontId || 0]
        sz = item.sz || font.sz || fonts[0].sz
        if (col[:longest].scan(/./mu).size * col[:sz]) < (item.value.to_s.scan(/./mu).size * sz)
          col[:sz] =  sz
          col[:longest] = item.value.to_s
        end
      end
      cells
    end

    # Determines the proper width for a column based on content.
    # @note
    #   width = Truncate([!{Number of Characters} * !{Maximum Digit Width} + !{5 pixel padding}]/!{Maximum Digit Width}*256)/256
    # @return [Float]
    # @param [Hash] A hash of auto_fit_data
    def auto_width(col)
      return col[:fixed] unless col[:fixed] == nil
      return Axlsx::FIXED_COL_WIDTH unless self.workbook.use_autowidth
      mdw_count, font_scale, mdw = 0, col[:sz]/11.0, 6.0
      mdw_count = col[:longest].scan(/./mu).reduce(0) do | count, char |
        count +=1 if @magick_draw.get_type_metrics(char).max_advance >= mdw
        count
      end
      ((mdw_count * mdw + 5) / mdw * 256) / 256.0 * font_scale
    end

    # Something to look into:
    #  width calculation actually needs to be done agains the formatted value for items that apply a
    # format
    # def excel_format(cell)
    #   # The most common case.
    #   return time.value.to_s if cell.style == 0
    #
    #   # The second most common case
    #   num_fmt = workbook.styles.cellXfs[items.style].numFmtId
    #   return value.to_s if num_fmt == 0
    #
    #   format_code = workbook.styles.numFmts[num_fmt]
    #   # need to find some exceptionally fast way of parsing value according to
    #   # an excel format_code
    #   item.value.to_s
    # end

  end
end
