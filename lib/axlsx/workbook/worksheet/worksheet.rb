# encoding: UTF-8
module Axlsx

  # The Worksheet class represents a worksheet in the workbook.
  class Worksheet

    # The name of the worksheet
    # @return [String]
    attr_reader :name

    # The sheet protection object for this workbook
    # @return [SheetProtection]
    # @see SheetProtection
    def sheet_protection
      @sheet_protection ||= SheetProtection.new
      yield @sheet_protection if block_given?
      @sheet_protection
    end

    # A collection of protected ranges in the worksheet
    # @note The recommended way to manage protected ranges is with Worksheet#protect_range
    # @see Worksheet#protect_range
    # @return [SimpleTypedList] The protected ranges for this worksheet
    attr_reader :protected_ranges

    # The sheet view object for this worksheet
    # @return [SheetView]
    # @see [SheetView]
    def sheet_view
      @sheet_view ||= SheetView.new
      yield @sheet_view if block_given?
      @sheet_view
    end

    # The workbook that owns this worksheet
    # @return [Workbook]
    attr_reader :workbook

    # The tables in this worksheet
    # @return [Array] of Table
    attr_reader :tables

    # The comments associated with this worksheet
    # @return [SimpleTypedList]
    attr_reader :comments

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
    # @deprecated Use {SheetView#show_grid_lines} instead.
    def show_gridlines
      warn('axlsx::DEPRECIATED: Worksheet#show_gridlines has been depreciated. This value can get over SheetView#show_grid_lines.')
      sheet_view.show_grid_lines
    end

    # Indicates if the worksheet is selected in the workbook
    # It is possible to have more than one worksheet selected, however it might cause issues
    # in some older versions of excel when using copy and paste.
    # @return Boolean
    # @deprecated Use {SheetView#tab_selected} instead.
    def selected
      warn('axlsx::DEPRECIATED: Worksheet#selected has been depreciated. This value can get over SheetView#tab_selected.')
      sheet_view.tab_selected
    end

    # Indicates if the worksheet will be fit by witdh or height to a specific number of pages.
    # To alter the width or height for page fitting, please use page_setup.fit_to_widht or page_setup.fit_to_height.
    # If you want the worksheet to fit on more pages (e.g. 2x2), set {PageSetup#fit_to_width} and {PageSetup#fit_to_height} accordingly.
    # @return Boolean
    # @see #page_setup
    def fit_to_page?
      return false unless @page_setup
      @page_setup.fit_to_page?
    end


    # Column info for the sheet
    # @return [SimpleTypedList]
    attr_reader :column_info

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

    # Page setup settings for printing the worksheet.
    # @example
    #      wb = Axlsx::Package.new.workbook
    #
    #      # using options when creating the worksheet.
    #      ws = wb.add_worksheet :page_setup => {:fit_to_width => 2, :orientation => :landscape}
    #
    #      # use the set method of the page_setup object
    #      ws.page_setup.set(:paper_width => "297mm", :paper_height => "210mm")
    #
    #      # setup page in a block
    #      ws.page_setup do |page|
    #        page.scale = 80
    #        page.orientation = :portrait
    #      end
    # @see PageSetup#initialize
    # @return [PageSetup]
    def page_setup
      @page_setup ||= PageSetup.new
      yield @page_setup if block_given?
      @page_setup
    end

    # Options for printing the worksheet.
    # @example
    #      wb = Axlsx::Package.new.workbook
    #      # using options when creating the worksheet.
    #      ws = wb.add_worksheet :print_options => {:grid_lines => true, :horizontal_centered => true}
    #
    #      # use the set method of the page_margins object
    #      ws.print_options.set(:headings => true)
    #
    #      # set page margins in a block
    #      ws.print_options do |options|
    #        options.horizontal_centered = true
    #        options.vertical_centered = true
    #      end
    # @see PrintOptions#initialize
    # @return [PrintOptions]
    def print_options
      @print_options ||= PrintOptions.new
      yield @print_options if block_given?
      @print_options
    end

    # definition of characters which are less than the maximum width of 0-9 in the default font for use in String#count.
    # This is used for autowidth calculations
    # @return [String]
    def self.thin_chars
      @thin_chars ||= "^.acefijklrstxyzFIJL()-"
    end

    # Creates a new worksheet.
    # @note the recommended way to manage worksheets is Workbook#add_worksheet
    # @see Workbook#add_worksheet
    # @option options [String] name The name of this worksheet.
    # @option options [Hash] page_margins A hash containing page margins for this worksheet. @see PageMargins
    # @option options [Hash] print_options A hash containing print options for this worksheet. @see PrintOptions
    # @option options [Boolean] show_gridlines indicates if gridlines should be shown for this sheet.
    def initialize(wb, options={})
      self.workbook = wb
      @workbook.worksheets << self
      @page_marging = @page_setup = @print_options = nil
      @drawing = @page_margins = @auto_filter = @sheet_protection = @sheet_view = nil
      @merged_cells = []
      @auto_fit_data = []
      @conditional_formattings = []
      @data_validations = []
      @comments = Comments.new(self)
      self.name = "Sheet" + (index+1).to_s
      @page_margins = PageMargins.new options[:page_margins] if options[:page_margins]
      @page_setup = PageSetup.new options[:page_setup]  if options[:page_setup]
      @print_options = PrintOptions.new options[:print_options] if options[:print_options]
      @rows = SimpleTypedList.new Row
      @column_info = SimpleTypedList.new Col
      @protected_ranges = SimpleTypedList.new ProtectedRange
      @tables = SimpleTypedList.new Table

      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
    end

    # convinience method to access all cells in this worksheet
    # @return [Array] cells
    def cells
      rows.flatten
    end

    # Add conditional formatting to this worksheet.
    #
    # @param [String] cells The range to apply the formatting to
    # @param [Array|Hash] rules An array of hashes (or just one) to create Conditional formatting rules from.
    # @example This would format column A whenever it is FALSE.
    #        # for a longer example, see examples/example_conditional_formatting.rb (link below)
    #        worksheet.add_conditional_formatting( "A1:A1048576", { :type => :cellIs, :operator => :equal, :formula => "FALSE", :dxfId => 1, :priority => 1 }
    #
    # @see ConditionalFormattingRule#initialize
    # @see file:examples/example_conditional_formatting.rb
    def add_conditional_formatting(cells, rules)
      cf = ConditionalFormatting.new( :sqref => cells )
      cf.add_rules rules
      @conditional_formattings << cf
    end

    # Add data validation to this worksheet.
    #
    # @param [String] cells The cells the validation will apply to.
    # @param [hash] data_validation options defining the validation to apply.
    # @see  examples/data_validation.rb for an example
    def add_data_validation(cells, data_validation)
      dv = DataValidation.new(data_validation)
      dv.sqref = cells
      @data_validations << dv
    end

    # Creates merge information for this worksheet.
    # Cells can be merged by calling the merge_cells method on a worksheet.
    # @example This would merge the three cells C1..E1    #
    #        worksheet.merge_cells "C1:E1"
    #        # you can also provide an array of cells to be merged
    #        worksheet.merge_cells worksheet.rows.first.cells[(2..4)]
    #        #alternatively you can do it from a single cell
    #        worksheet["C1"].merge worksheet["E1"]
    # @param [Array, string] cells
    def merge_cells(cells)
      @merged_cells << if cells.is_a?(String)
                         cells
      elsif cells.is_a?(Array)
        Axlsx::cell_range(cells, false)
      end
    end
    # Adds a new protected cell range to the worksheet. Note that protected ranges are only in effect when sheet protection is enabled.
    # @param [String|Array] cells The string reference for the cells to protect or an array of cells.
    # @return [ProtectedRange]
    # @note When using an array of cells, a contiguous range is created from the minimum top left to the maximum top bottom of the cells provided.
    def protect_range(cells)
      sqref = if cells.is_a?(String)
                cells
              elsif cells.is_a?(SimpleTypedList)
                Axlsx::cell_range(cells, false)
              end
      @protected_ranges << ProtectedRange.new(:sqref => sqref, :name => 'Range#{@protected_ranges.size}')
      @protected_ranges.last
    end

    # The demensions of a worksheet. This is not actually a required element by the spec,
    # but at least a few other document readers expect this for conversion
    # @return [String] the A1:B2 style reference for the first and last row column intersection in the workbook
    def dimension
      "#{dimension_reference(rows.first.cells.first, 'A1')}:#{dimension_reference(rows.last.cells.last, 'AA200')}"
    end

    #
    # Indicates if gridlines should be shown in the sheet.
    # This is true by default.
    # @return [Boolean]
    # @deprecated Use {SheetView#show_grid_lines=} instead.
    def show_gridlines=(v)
      warn('axlsx::DEPRECIATED: Worksheet#show_gridlines= has been depreciated. This value can be set over SheetView#show_grid_lines=.')
      Axlsx::validate_boolean v
      sheet_view.show_grid_lines = v
    end

    # @see selected
    # @return [Boolean]
    # @deprecated Use {SheetView#tab_selected=} instead.
    def selected=(v)
      warn('axlsx::DEPRECIATED: Worksheet#selected= has been depreciated. This value can be set over SheetView#tab_selected=.')
      Axlsx::validate_boolean v
      sheet_view.tab_selected = v
    end


    # (see #fit_to_page)
    # @return [Boolean]
    def fit_to_page=(v)
      warn('axlsx::DEPRECIATED: Worksheet#fit_to_page has been depreciated. This value will automatically be set for you when you use PageSetup#fit_to.')
      fit_to_page?
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
      raise ArgumentError, (ERR_SHEET_NAME_COLON_FORBIDDEN % v) if v.include? ':'
      v = Axlsx::coder.encode(v) 
      sheet_names = @workbook.worksheets.map { |s| s.name }
      raise ArgumentError, (ERR_DUPLICATE_SHEET_NAME % v) if sheet_names.include?(v)
      @name=v
    end

    # The absolute auto filter range
    # @see auto_filter
    def abs_auto_filter
      Axlsx.cell_range(@auto_filter.split(':').collect { |name| name_to_cell(name)}) if @auto_filter
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
      update_column_info @rows.last.cells, options.delete(:widths) || []
      yield @rows.last if block_given?
      @rows.last
    end

    alias :<< :add_row

    # Set the style for cells in a specific row
    # @param [Integer] index or range of indexes in the table
    # @param [Integer] style the cellXfs index
    # @param [Hash] options the options used when applying the style
    # @option [Integer] :col_offset only cells after this column will be updated.
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
    # @param [Integer] style the cellXfs index
    # @param [Hash] options
    # @option [Integer] :row_offset only cells after this column will be updated.
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
    # Setting a fixed column width to nil will revert the behaviour back to calculating the width for you on the next call to add_row.
    # @example This would set the first and third column widhts but leave the second column in autofit state.
    #      ws.column_widths 7.2, nil, 3
    # @note For updating only a single column it is probably easier to just set the width of the ws.column_info[col_index].width directly
    # @param [Integer|Float|Fixnum|nil] widths
    def column_widths(*widths)
      widths.each_with_index do |value, index|
        next if value == nil
        Axlsx::validate_unsigned_numeric(value) unless value == nil
        @column_info[index] ||= Col.new index+1, index+1
        @column_info[index].width = value
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


    # Shortcut to comments#add_comment
    def add_comment(options={})
      @comments.add_comment(options)
    end

    # Adds a media item to the worksheets drawing
    # @option [Hash] options options passed to drawing.add_image
    def add_image(options={})
      image = drawing.add_image(options)
      yield image if block_given?
      image
    end

    # Serializes the worksheet object to an xml string
    # This intentionally does not use nokogiri for performance reasons
    # @return [String]
    def to_xml_string
      str = '<?xml version="1.0" encoding="UTF-8"?>'
      str << worksheet_node
      str << sheet_pr_node
      str << dimension_node
      @sheet_view.to_xml_string(str) if @sheet_view
      str << cols_node
      str << sheet_data_node

      str << auto_filter_node
      @sheet_protection.to_xml_string(str) if @sheet_protection
      str << protected_ranges_node
      str << merged_cells_node
      @print_options.to_xml_string(str) if @print_options
      page_margins.to_xml_string(str) if @page_margins
      page_setup.to_xml_string(str) if @page_setup
      str << drawing_node
      str << legacy_drawing_node
      str << table_parts_node
      str << conditional_formattings_node
      str << data_validations_node 
      str << '</worksheet>'
      # User reported that when parsing some old data that had control characters excel chokes.
      # All of the following are defined as illegal xml characters in the xml spec, but for now I am only dealing with control 
      # characters. Thanks to asakusarb and @hsbt's flash of code on the screen!
      # [#x1-#x8], [#xB-#xC], [#xE-#x1F], [#x7F-#x84], [#x86-#x9F], [#xFDD0-#xFDDF],
      # [#x1FFFE-#x1FFFF], [#x2FFFE-#x2FFFF], [#x3FFFE-#x3FFFF],
      # [#x4FFFE-#x4FFFF], [#x5FFFE-#x5FFFF], [#x6FFFE-#x6FFFF],
      # [#x7FFFE-#x7FFFF], [#x8FFFE-#x8FFFF], [#x9FFFE-#x9FFFF],
      # [#xAFFFE-#xAFFFF], [#xBFFFE-#xBFFFF], [#xCFFFE-#xCFFFF],
      # [#xDFFFE-#xDFFFF], [#xEFFFE-#xEFFFF], [#xFFFFE-#xFFFFF],
      # [#x10FFFE-#x10FFFF].
      str.gsub(/[[:cntrl:]]/,'')
    end

    # The worksheet relationships. This is managed automatically by the worksheet
    # @return [Relationships]
    def relationships
      r = Relationships.new
      @tables.each do |table|
        r << Relationship.new(TABLE_R, "../#{table.pn}")
      end

      r << Relationship.new(VML_DRAWING_R, "../#{@comments.vml_drawing.pn}") if @comments.size > 0
      r << Relationship.new(COMMENT_R, "../#{@comments.pn}") if @comments.size > 0
      r << Relationship.new(COMMENT_R_NULL, "NULL") if @comments.size > 0

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

    # Helper method for parsingout the root node for worksheet
    # @return [String]
    def worksheet_node
      "<worksheet xmlns=\"%s\" xmlns:r=\"%s\">" % [XML_NS, XML_NS_R]
    end

    # Helper method fo parsing out the sheetPr node
    # @return [String]
    def sheet_pr_node
      return '' unless fit_to_page?
      "<sheetPr><pageSetUpPr fitToPage=\"%s\"></pageSetUpPr></sheetPr>" % fit_to_page?
    end

    # Helper method for parsing out the demension node
    # @return [String]
    def dimension_node
      return '' if rows.size == 0
      "<dimension ref=\"%s\"></dimension>" % dimension
    end

    # Helper method for parsing out the sheetData node
    # @return [String]
    def sheet_data_node
      str = '<sheetData>'
      @rows.each_with_index { |row, index| row.to_xml_string(index, str) }
      str << '</sheetData>'
    end

    # Helper method for parsing out the autoFilter node
    # @return [String]
    def auto_filter_node
      return '' unless @auto_filter
      "<autoFilter ref='%s'></autoFilter>" % @auto_filter
    end

    # Helper method for parsing out the cols node
    # @return [String]
    def cols_node
      return '' if @column_info.empty?
      str = "<cols>"
      @column_info.each { |col| col.to_xml_string(str) }
      str << '</cols>'
    end 

    # Helper method for parsing out the protectedRanges node
    # @return [String]
    def protected_ranges_node
      return '' if @protected_ranges.empty?
      str = '<protectedRanges>'
      @protected_ranges.each { |pr| pr.to_xml_string(str) }
      str << '</protectedRanges>'
    end

    # Helper method for parsing out the mergedCells node
    # @return [String]
    def merged_cells_node
      return '' if @merged_cells.size == 0
      str = "<mergeCells count='#{@merged_cells.size}'>"
      @merged_cells.each { |merged_cell| str << "<mergeCell ref='#{merged_cell}'></mergeCell>" }
      str << '</mergeCells>'
    end

    # Helper method for parsing out the drawing node
    # @return [String]
    def drawing_node
      return '' unless @drawing
      "<drawing r:id='rId" << (relationships.index{ |r| r.Type == DRAWING_R } + 1).to_s << "'/>" 
    end

    # Helper method for parsing out the legacyDrawing node required for comments
    # @return [String]
    def legacy_drawing_node
      return '' if @comments.empty?
      "<legacyDrawing r:id='rId" << (relationships.index{ |r| r.Type == VML_DRAWING_R } + 1).to_s << "'/>" 
    end

    # Helper method for parsing out the tableParts node
    # @return [String]
    def table_parts_node
      return '' if @tables.empty?
      str = "<tableParts count='#{@tables.size}'>"
      @tables.each { |table| str << "<tablePart r:id='#{table.rId}'/>" }
      str << '</tableParts>'
    end

    # Helper method for parsing out the conditional formattings
    # @return [String]
    def conditional_formattings_node
      return '' if @conditional_formattings.size == 0
      str = ''
      @conditional_formattings.each { |conditional_formatting| str << conditional_formatting.to_xml_string }
      str
    end

    # Helper method for parsing out the dataValidations node
    # @return [String]
    def data_validations_node
      return '' if @data_validations.size == 0
      str = "<dataValidations count='#{@data_validations.size}'>"
      @data_validations.each { |data_validation| str << data_validation.to_xml_string }
      str << '</dataValidations>'
    end

    # assigns the owner workbook for this worksheet
    def workbook=(v) DataTypeValidator.validate "Worksheet.workbook", Workbook, v; @workbook = v; end

    def styles
      @styles ||= self.workbook.styles
    end

    def update_column_info(cells, widths=[])
      cells.each_with_index do |cell, index|
        col = find_or_create_column_info(index)
        next if widths[index] == :ignore
        col.update_width(cell, widths[index], workbook.use_autowidth)
      end
    end

    def find_or_create_column_info(index, fixed_width=nil)
      col = @column_info[index] || Col.new(index + 1, index + 1)
      @column_info[index] = col if index == @column_info.size
      col
    end

    def dimension_reference(cell, default)
      return default unless cell.respond_to?(:r)
      cell.r
    end
  end
end
