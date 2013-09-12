# encoding: UTF-8
module Axlsx

  # The Worksheet class represents a worksheet in the workbook.
  class Worksheet
    include Axlsx::OptionsParser

    # definition of characters which are less than the maximum width of 0-9 in the default font for use in String#count.
    # This is used for autowidth calculations
    # @return [String]
    def self.thin_chars
      # removed 'e' and 'y' from this list - as a GUESS
      @thin_chars ||= "^.acfijklrstxzFIJL()-"
    end

    # Creates a new worksheet.
    # @note the recommended way to manage worksheets is Workbook#add_worksheet
    # @see Workbook#add_worksheet
    # @option options [String] name The name of this worksheet.
    # @option options [Hash] page_margins A hash containing page margins for this worksheet. @see PageMargins
    # @option options [Hash] print_options A hash containing print options for this worksheet. @see PrintOptions
    # @option options [Hash] header_footer A hash containing header/footer options for this worksheet. @see HeaderFooter
    # @option options [Boolean] show_gridlines indicates if gridlines should be shown for this sheet.
    def initialize(wb, options={})
      self.workbook = wb
      @sheet_protection = nil

      initialize_page_options(options)
      parse_options options
      @workbook.worksheets << self
    end

    # Initalizes page margin, setup and print options
    # @param [Hash] options Options passed in from the initializer
    def initialize_page_options(options)
      @page_margins = PageMargins.new options[:page_margins] if options[:page_margins]
      @page_setup = PageSetup.new options[:page_setup]  if options[:page_setup]
      @print_options = PrintOptions.new options[:print_options] if options[:print_options]
      @header_footer = HeaderFooter.new options[:header_footer] if options[:header_footer]
      @row_breaks = RowBreaks.new
      @col_breaks = ColBreaks.new
    end

    # The name of the worksheet
    # @return [String]
    def name
      @name ||=  "Sheet" + (index+1).to_s
    end

    # The sheet calculation properties
    # @return [SheetCalcPr]
    def sheet_calc_pr
      @sheet_calc_pr ||= SheetCalcPr.new
    end

    # The sheet protection object for this workbook
    # @return [SheetProtection]
    # @see SheetProtection
    def sheet_protection
      @sheet_protection ||= SheetProtection.new
      yield @sheet_protection if block_given?
      @sheet_protection
    end

    # The sheet view object for this worksheet
    # @return [SheetView]
    # @see [SheetView]
    def sheet_view
      @sheet_view ||= SheetView.new
      yield @sheet_view if block_given?
      @sheet_view
    end

    # The sheet format pr for this worksheet
    # @return [SheetFormatPr]
    # @see [SheetFormatPr]
    def sheet_format_pr
      @sheet_format_pr ||= SheetFormatPr.new
      yeild @sheet_format_pr if block_given?
      @sheet_format_pr
    end

    # The workbook that owns this worksheet
    # @return [Workbook]
    attr_reader :workbook

    # The tables in this worksheet
    # @return [Array] of Table
    def tables
      @tables ||=  Tables.new self
    end

    # The pivot tables in this worksheet
    # @return [Array] of Table
    def pivot_tables
      @pivot_tables ||=  PivotTables.new self
    end

    # A collection of column breaks added to this worksheet
    # @note Please do not use this directly. Instead use
    # add_page_break
    # @see Worksheet#add_page_break
    def col_breaks
      @col_breaks ||= ColBreaks.new
    end

    # A collection of row breaks added to this worksheet
    # @note Please do not use this directly. Instead use
    # add_page_break
    # @see Worksheet#add_page_break
    def row_breaks
      @row_breaks ||= RowBreaks.new
    end

    # A typed collection of hyperlinks associated with this worksheet
    # @return [WorksheetHyperlinks]
    def hyperlinks
      @hyperlinks ||= WorksheetHyperlinks.new self
    end

    # The a shortcut to the worksheet_comments list of comments
    # @return [Array|SimpleTypedList]
    def comments
      worksheet_comments.comments if worksheet_comments.has_comments?
    end

    # The rows in this worksheet
    # @note The recommended way to manage rows is Worksheet#add_row
    # @return [SimpleTypedList]
    # @see Worksheet#add_row
    def rows
      @rows ||=  SimpleTypedList.new Row
    end

    # returns the sheet data as columns
    # If you pass a block, it will be evaluated whenever a row does not have a
    # cell at a specific index. The block will be called with the row and column
    # index in the missing cell was found.
    # @example
    #     cols { |row_index, column_index| p "warn - row #{row_index} is does not have a cell at #{column_index}
    def cols(&block)
      @rows.transpose(&block)
    end

    # An range that excel will apply an auto-filter to "A1:B3"
    # This will turn filtering on for the cells in the range.
    # The first row is considered the header, while subsequent rows are considered to be data.
    # @return String
    def auto_filter
      @auto_filter ||= AutoFilter.new self
    end

    # Indicates if the worksheet will be fit by witdh or height to a specific number of pages.
    # To alter the width or height for page fitting, please use page_setup.fit_to_widht or page_setup.fit_to_height.
    # If you want the worksheet to fit on more pages (e.g. 2x2), set {PageSetup#fit_to_width} and {PageSetup#fit_to_height} accordingly.
    # @return Boolean
    # @see #page_setup
    def fit_to_page?
      return false unless self.instance_values.keys.include?('page_setup')
      page_setup.fit_to_page?
    end


    # Column info for the sheet
    # @return [SimpleTypedList]
    def column_info
      @column_info ||= Cols.new self
    end

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

    # Options for headers and footers.
    # @example
    #   wb = Axlsx::Package.new.workbook
    #   # would generate something like: "file.xlsx : sheet_name 2 of 7 date with timestamp"
    #   header = {:different_odd_ => false, :odd_header => "&L&F : &A&C&Pof%N%R%D %T"}
    #   ws = wb.add_worksheet :header_footer => header
    #
    # @see HeaderFooter#initialize
    # @return [HeaderFooter]
    def header_footer
      @header_footer ||= HeaderFooter.new
      yield @header_footer if block_given?
      @header_footer
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
    # @param [Array, string] cells
    def merge_cells(cells)
      merged_cells.add cells
    end

    # Adds a new protected cell range to the worksheet. Note that protected ranges are only in effect when sheet protection is enabled.
    # @param [String|Array] cells The string reference for the cells to protect or an array of cells.
    # @return [ProtectedRange]
    # @note When using an array of cells, a contiguous range is created from the minimum top left to the maximum top bottom of the cells provided.
    def protect_range(cells)
      protected_ranges.add_range(cells)
    end

    # The dimensions of a worksheet. This is not actually a required element by the spec,
    # but at least a few other document readers expect this for conversion
    # @return [Dimension]
    def dimension
      @dimension ||= Dimension.new self
    end

    # The sheet properties for this workbook.
    # Currently only pageSetUpPr -> fitToPage is implemented
    # @return [SheetPr]
    def sheet_pr
      @sheet_pr ||= SheetPr.new self
    end

    # Indicates if gridlines should be shown in the sheet.
    # This is true by default.
    # @return [Boolean]
    # @deprecated Use SheetView#show_grid_lines= instead.
    def show_gridlines=(v)
      warn('axlsx::DEPRECIATED: Worksheet#show_gridlines= has been depreciated. This value can be set over SheetView#show_grid_lines=.')
      Axlsx::validate_boolean v
      sheet_view.show_grid_lines = v
    end

    # @see selected
    # @return [Boolean]
    # @deprecated Use SheetView#tab_selected= instead.
    def selected=(v)
      warn('axlsx::DEPRECIATED: Worksheet#selected= has been depreciated. This value can be set over SheetView#tab_selected=.')
      Axlsx::validate_boolean v
      sheet_view.tab_selected = v
    end

    # Indicates if the worksheet should show gridlines or not
    # @return Boolean
    # @deprecated Use SheetView#show_grid_lines instead.
    def show_gridlines
      warn('axlsx::DEPRECIATED: Worksheet#show_gridlines has been depreciated. This value can get over SheetView#show_grid_lines.')
      sheet_view.show_grid_lines
    end

    # Indicates if the worksheet is selected in the workbook
    # It is possible to have more than one worksheet selected, however it might cause issues
    # in some older versions of excel when using copy and paste.
    # @return Boolean
    # @deprecated Use SheetView#tab_selected instead.
    def selected
      warn('axlsx::DEPRECIATED: Worksheet#selected has been depreciated. This value can get over SheetView#tab_selected.')
      sheet_view.tab_selected
    end

    # (see #fit_to_page)
    # @return [Boolean]
    def fit_to_page=(v)
      warn('axlsx::DEPRECIATED: Worksheet#fit_to_page has been depreciated. This value will automatically be set for you when you use PageSetup#fit_to.')
      fit_to_page?
    end

    # The name of the worksheet
    # The name of a worksheet must be unique in the workbook, and must not exceed 31 characters
    # @param [String] name
    def name=(name)
      validate_sheet_name name
      @name=Axlsx::coder.encode(name)
    end

    # The auto filter range for the worksheet
    # @param [String] v
    # @see auto_filter
    def auto_filter=(v)
      DataTypeValidator.validate "Worksheet.auto_filter", String, v
      auto_filter.range = v
    end

    # Accessor for controlling whether leading and trailing spaces in cells are
    # preserved or ignored. The default is to preserve spaces.
    attr_accessor :preserve_spaces

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

    # The relationship id of this worksheet.
    # @return [String]
    # @see Relationship#Id
    def rId
      @workbook.relationships.for(self).Id
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
      worksheet_drawing.drawing
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
      conditional_formattings << cf
      conditional_formattings
    end

    # Add data validation to this worksheet.
    #
    # @param [String] cells The cells the validation will apply to.
    # @param [hash] data_validation options defining the validation to apply.
    # @see  examples/data_validation.rb for an example
    def add_data_validation(cells, data_validation)
      dv = DataValidation.new(data_validation)
      dv.sqref = cells
      data_validations << dv
    end

    # Adds a new hyperlink to the worksheet
    # @param [Hash] options for the hyperlink
    # @see WorksheetHyperlink for a list of options
    # @return [WorksheetHyperlink]
    def add_hyperlink(options={})
      hyperlinks.add(options)
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
      chart = worksheet_drawing.add_chart(chart_type, options)
      yield chart if block_given?
      chart
    end

    # needs documentation
    def add_table(ref, options={})
      tables << Table.new(ref, self, options)
      yield tables.last if block_given?
      tables.last
    end

    def add_pivot_table(ref, range, options={})
      pivot_tables << PivotTable.new(ref, range, self, options)
      yield pivot_tables.last if block_given?
      pivot_tables.last
    end

    # Shortcut to worsksheet_comments#add_comment
    def add_comment(options={})
      worksheet_comments.add_comment(options)
    end

    # Adds a media item to the worksheets drawing
    # @option [Hash] options options passed to drawing.add_image
    def add_image(options={})
      image = worksheet_drawing.add_image(options)
      yield image if block_given?
      image
    end

    # Adds a page break (row break) to the worksheet
    # @param cell A Cell object or excel style string reference indicating where the break
    # should be added to the sheet.
    # @example
    #   ws.add_page_break("A4")
    def add_page_break(cell)
      DataTypeValidator.validate "Worksheet#add_page_break cell", [String, Cell], cell
      column_index, row_index = if cell.is_a?(String)
          Axlsx.name_to_indices(cell)
        else
          cell.pos
        end
      if column_index > 0
        col_breaks.add_break(:id => column_index)
      end
      row_breaks.add_break(:id => row_index)
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
        find_or_create_column_info(index).width = value
      end
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
      cells = @rows[(offset..-1)].map { |row| row.cells[index] }.flatten.compact
      cells.each { |cell| cell.style = style }
    end

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
      cells = cols[(offset..-1)].map { |column| column[index] }.flatten.compact
      cells.each { |cell| cell.style = style }
    end

    # Serializes the worksheet object to an xml string
    # This intentionally does not use nokogiri for performance reasons
    # @return [String]
    def to_xml_string
      auto_filter.apply if auto_filter.range
      str = '<?xml version="1.0" encoding="UTF-8"?>'
      str << worksheet_node
      serializable_parts.each do |item|
        item.to_xml_string(str) if item
      end
      str << '</worksheet>'
      Axlsx::sanitize(str)
    end

    # The worksheet relationships. This is managed automatically by the worksheet
    # @return [Relationships]
    def relationships
      r = Relationships.new
      r + [tables.relationships,
           worksheet_comments.relationships,
           hyperlinks.relationships,
           worksheet_drawing.relationship,
           pivot_tables.relationships].flatten.compact || []
      r
    end

    # Returns the cell or cells defined using excel style A1:B3 references.
    # @param [String|Integer] cell_def the string defining the cell or range of cells, or the rownumber
    # @return [Cell, Array]
    def [] (cell_def)
      return rows[cell_def] if cell_def.is_a?(Integer)
      parts = cell_def.split(':').map{ |part| name_to_cell part }
      if parts.size == 1
        parts.first
      else
        range(*parts)
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

    # shortcut method to access styles direclty from the worksheet
    # This lets us do stuff like:
    # @example
    #     p = Axlsx::Package.new
    #     p.workbook.add_worksheet(:name => 'foo') do |sheet|
    #       my_style = sheet.styles.add_style { :bg_color => "FF0000" }
    #       sheet.add_row ['Oh No!'], :styles => my_style
    #     end
    #     p.serialize 'foo.xlsx'
    def styles
      @styles ||= self.workbook.styles
    end

    # shortcut level to specify the outline level for a series of rows
    # Oulining is what lets you add collapse and expand to a data set.
    # @param [Integer] start_index The zero based index of the first row of outlining.
    # @param [Integer] end_index The zero based index of  the last row to be outlined
    # @param [integer] level The level of outline to apply
    # @param [Integer] collapsed The initial collapsed state of the outline group
    def outline_level_rows(start_index, end_index, level = 1, collapsed = true)
      outline rows, (start_index..end_index), level, collapsed
    end

    # shortcut level to specify the outline level for a series of columns
    # Oulining is what lets you add collapse and expand to a data set.
    # @param [Integer] start_index The zero based index of the first column of outlining.
    # @param [Integer] end_index The zero based index of  the last column to be outlined
    # @param [integer] level The level of outline to apply
    # @param [Integer] collapsed The initial collapsed state of the outline group
    def outline_level_columns(start_index, end_index, level = 1, collapsed = true)
      outline column_info, (start_index..end_index), level, collapsed
    end

    private

    def xml_space
      workbook.xml_space
    end

    def outline(collection, range, level = 1, collapsed = true)
       range.each do |index|
        unless (item = collection[index]).nil?
          item.outline_level = level
          item.hidden = collapsed
        end
        sheet_view.show_outline_symbols = true
      end
    end

    def validate_sheet_name(name)
      DataTypeValidator.validate "Worksheet.name", String, name
      raise ArgumentError, (ERR_SHEET_NAME_TOO_LONG % name) if name.size > 31
      raise ArgumentError, (ERR_SHEET_NAME_CHARACTER_FORBIDDEN % name) if '[]*/\?:'.chars.any? { |char| name.include? char }
      name = Axlsx::coder.encode(name)
      sheet_names = @workbook.worksheets.reject { |s| s == self }.map { |s| s.name }
      raise ArgumentError, (ERR_DUPLICATE_SHEET_NAME % name) if sheet_names.include?(name)
    end

    def serializable_parts
      [sheet_pr, dimension, sheet_view, sheet_format_pr, column_info,
       sheet_data, sheet_calc_pr, @sheet_protection, protected_ranges,
       auto_filter, merged_cells, conditional_formattings,
       data_validations, hyperlinks, print_options, page_margins,
       page_setup, header_footer, row_breaks, col_breaks, worksheet_drawing, worksheet_comments,
       tables]
    end

    def range(*cell_def)
      first, last = cell_def
      cells = []
      rows[(first.row.index..last.row.index)].each do |r|
        r.cells[(first.index..last.index)].each do |c|
          cells << c
        end
      end
      cells
    end

    # A collection of protected ranges in the worksheet
    # @note The recommended way to manage protected ranges is with Worksheet#protect_range
    # @see Worksheet#protect_range
    # @return [SimpleTypedList] The protected ranges for this worksheet
    def protected_ranges
      @protected_ranges ||= ProtectedRanges.new self
      # SimpleTypedList.new ProtectedRange
    end

    # conditional formattings
    # @return [Array]
    def conditional_formattings
      @conditional_formattings ||= ConditionalFormattings.new self
    end

    # data validations array
    # @return [Array]
    def data_validations
      @data_validations ||= DataValidations.new self
    end

    # merged cells array
    # @return [Array]
    def merged_cells
      @merged_cells ||= MergedCells.new self
    end


    # Helper method for parsingout the root node for worksheet
    # @return [String]
    def worksheet_node
       "<worksheet xmlns=\"%s\" xmlns:r=\"%s\" xml:space=\"#{xml_space}\">" % [XML_NS, XML_NS_R]
    end

    def sheet_data
      @sheet_data ||= SheetData.new self
    end

    def worksheet_drawing
      @worksheet_drawing ||= WorksheetDrawing.new self
    end

    # The comments associated with this worksheet
    # @return [SimpleTypedList]
    def worksheet_comments
      @worksheet_comments ||= WorksheetComments.new self
    end

    def workbook=(v) DataTypeValidator.validate "Worksheet.workbook", Workbook, v; @workbook = v; end

    def update_column_info(cells, widths=[])
      cells.each_with_index do |cell, index|
        col = find_or_create_column_info(index)
        next if widths[index] == :ignore
        col.update_width(cell, widths[index], workbook.use_autowidth)
      end
    end

    def find_or_create_column_info(index)
      column_info[index] ||= Col.new(index + 1, index + 1)
    end

  end
end
