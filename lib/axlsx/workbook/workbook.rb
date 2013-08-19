# -*- coding: utf-8 -*-
module Axlsx
require 'axlsx/workbook/worksheet/sheet_calc_pr.rb'
require 'axlsx/workbook/worksheet/auto_filter/auto_filter.rb'
require 'axlsx/workbook/worksheet/date_time_converter.rb'
require 'axlsx/workbook/worksheet/protected_range.rb'
require 'axlsx/workbook/worksheet/protected_ranges.rb'
require 'axlsx/workbook/worksheet/cell_serializer.rb'
require 'axlsx/workbook/worksheet/cell.rb'
require 'axlsx/workbook/worksheet/page_margins.rb'
require 'axlsx/workbook/worksheet/page_set_up_pr.rb'
require 'axlsx/workbook/worksheet/page_setup.rb'
require 'axlsx/workbook/worksheet/header_footer.rb'
require 'axlsx/workbook/worksheet/print_options.rb'
require 'axlsx/workbook/worksheet/cfvo.rb'
require 'axlsx/workbook/worksheet/cfvos.rb'
require 'axlsx/workbook/worksheet/color_scale.rb'
require 'axlsx/workbook/worksheet/data_bar.rb'
require 'axlsx/workbook/worksheet/icon_set.rb'
require 'axlsx/workbook/worksheet/conditional_formatting.rb'
require 'axlsx/workbook/worksheet/conditional_formatting_rule.rb'
require 'axlsx/workbook/worksheet/conditional_formattings.rb'
require 'axlsx/workbook/worksheet/row.rb'
require 'axlsx/workbook/worksheet/col.rb'
require 'axlsx/workbook/worksheet/cols.rb'
require 'axlsx/workbook/worksheet/comments.rb'
require 'axlsx/workbook/worksheet/comment.rb'
require 'axlsx/workbook/worksheet/merged_cells.rb'
require 'axlsx/workbook/worksheet/sheet_protection.rb'
require 'axlsx/workbook/worksheet/sheet_pr.rb'
require 'axlsx/workbook/worksheet/dimension.rb'
require 'axlsx/workbook/worksheet/sheet_data.rb'
require 'axlsx/workbook/worksheet/worksheet_drawing.rb'
require 'axlsx/workbook/worksheet/worksheet_comments.rb'
require 'axlsx/workbook/worksheet/worksheet_hyperlink'
require 'axlsx/workbook/worksheet/worksheet_hyperlinks'
require 'axlsx/workbook/worksheet/break'
require 'axlsx/workbook/worksheet/row_breaks'
require 'axlsx/workbook/worksheet/col_breaks'



require 'axlsx/workbook/worksheet/worksheet.rb'
require 'axlsx/workbook/shared_strings_table.rb'
require 'axlsx/workbook/defined_name.rb'
require 'axlsx/workbook/defined_names.rb'
require 'axlsx/workbook/worksheet/table_style_info.rb'
require 'axlsx/workbook/worksheet/table.rb'
require 'axlsx/workbook/worksheet/tables.rb'
require 'axlsx/workbook/worksheet/pivot_table_cache_definition.rb'
require 'axlsx/workbook/worksheet/pivot_table.rb'
require 'axlsx/workbook/worksheet/pivot_tables.rb'
require 'axlsx/workbook/worksheet/data_validation.rb'
require 'axlsx/workbook/worksheet/data_validations.rb'
require 'axlsx/workbook/worksheet/sheet_view.rb'
require 'axlsx/workbook/worksheet/sheet_format_pr.rb'
require 'axlsx/workbook/worksheet/pane.rb'
require 'axlsx/workbook/worksheet/selection.rb'
  # The Workbook class is an xlsx workbook that manages worksheets, charts, drawings and styles.
  # The following parts of the Office Open XML spreadsheet specification are not implimented in this version.
  #
  #   bookViews
  #   calcPr
  #   customWorkbookViews
  #   definedNames
  #   externalReferences
  #   extLst
  #   fileRecoveryPr
  #   fileSharing
  #   fileVersion
  #   functionGroups
  #   oleSize
  #   pivotCaches
  #   smartTagPr
  #   smartTagTypes
  #   webPublishing
  #   webPublishObjects
  #   workbookProtection
  #   workbookPr*
  #
  #   *workbookPr is only supported to the extend of date1904
  class Workbook

    # When true, the Package will be generated with a shared string table. This may be required by some OOXML processors that do not
    # adhere to the ECMA specification that dictates string may be inline in the sheet.
    # Using this option will increase the time required to serialize the document as every string in every cell must be analzed and referenced.
    # @return [Boolean]
    attr_reader :use_shared_strings

    # @see use_shared_strings
    def use_shared_strings=(v)
      Axlsx::validate_boolean(v)
      @use_shared_strings = v
    end


   # A collection of worksheets associated with this workbook.
    # @note The recommended way to manage worksheets is add_worksheet
    # @see Workbook#add_worksheet
    # @see Worksheet
    # @return [SimpleTypedList]
    attr_reader :worksheets

    # A colllection of charts associated with this workbook
    # @note The recommended way to manage charts is Worksheet#add_chart
    # @see Worksheet#add_chart
    # @see Chart
    # @return [SimpleTypedList]
    attr_reader :charts

    # A colllection of images associated with this workbook
    # @note The recommended way to manage images is Worksheet#add_image
    # @see Worksheet#add_image
    # @see Pic
    # @return [SimpleTypedList]
    attr_reader :images

    # A colllection of drawings associated with this workbook
    # @note The recommended way to manage drawings is Worksheet#add_chart
    # @see Worksheet#add_chart
    # @see Drawing
    # @return [SimpleTypedList]
    attr_reader :drawings

    # pretty sure this two are always empty and can be removed.


    # A colllection of tables associated with this workbook
    # @note The recommended way to manage drawings is Worksheet#add_table
    # @see Worksheet#add_table
    # @see Table
    # @return [SimpleTypedList]
    attr_reader :tables

    # A colllection of pivot tables associated with this workbook
    # @note The recommended way to manage drawings is Worksheet#add_table
    # @see Worksheet#add_table
    # @see Table
    # @return [SimpleTypedList]
    attr_reader :pivot_tables


    # A collection of defined names for this workbook
    # @note The recommended way to manage defined names is Workbook#add_defined_name
    # @see DefinedName
    # @return [DefinedNames]
    def defined_names
      @defined_names ||= DefinedNames.new
    end

    # A collection of comments associated with this workbook
    # @note The recommended way to manage comments is WOrksheet#add_comment
    # @see Worksheet#add_comment
    # @see Comment
    # @return [Comments]
    def comments
      worksheets.map { |sheet| sheet.comments }.compact
    end

    # The styles associated with this workbook
    # @note The recommended way to manage styles is Styles#add_style
    # @see Style#add_style
    # @see Style
    # @return [Styles]
    def styles
      yield @styles if block_given?
      @styles
    end


    # Indicates if the epoc date for serialization should be 1904. If false, 1900 is used.
    @@date1904 = false


    # A quick helper to retrive a worksheet by name
    # @param [String] name The name of the sheet you are looking for
    # @return [Worksheet] The sheet found, or nil
    def sheet_by_name(name)
      index = @worksheets.index { |sheet| sheet.name == name }
      @worksheets[index] if index
    end

    # lets come back to this later when we are ready for parsing.
    #def self.parse entry
    #  io = entry.get_input_stream
    #  w = self.new
    #  w.parser_xml = Nokogiri::XML(io.read)
    #  w.parse_string :date1904, "//xmlns:workbookPr/@date1904"
    #  w
    #end

    # Creates a new Workbook
    # The recomended way to work with workbooks is via Package#workbook
    # @option options [Boolean] date1904. If this is not specified, date1904 is set to false. Office 2011 for Mac defaults to false.
    def initialize(options={})
      @styles = Styles.new
      @worksheets = SimpleTypedList.new Worksheet
      @drawings = SimpleTypedList.new Drawing
      @charts = SimpleTypedList.new Chart
      @images = SimpleTypedList.new Pic
      # Are these even used????? Check package serialization parts
      @tables = SimpleTypedList.new Table
      @pivot_tables = SimpleTypedList.new PivotTable
      @comments = SimpleTypedList.new Comments


      @use_autowidth = true

      self.date1904= !options[:date1904].nil? && options[:date1904]
      yield self if block_given?
    end

    # Instance level access to the class variable 1904
    # @return [Boolean]
    def date1904() @@date1904; end

    # see @date1904
    def date1904=(v) Axlsx::validate_boolean v; @@date1904 = v; end

    # Sets the date1904 attribute to the provided boolean
    # @return [Boolean]
    def self.date1904=(v) Axlsx::validate_boolean v; @@date1904 = v; end

    # retrieves the date1904 attribute
    # @return [Boolean]
    def self.date1904() @@date1904; end

    # Indicates if the workbook should use autowidths or not.
    # @note This gem no longer depends on RMagick for autowidth
    #     calculation. Thus the performance benefits of turning this off are
    #     marginal unless you are creating a very large sheet.
    # @return [Boolean]
    def use_autowidth() @use_autowidth; end

    # see @use_autowidth
    def use_autowidth=(v=true) Axlsx::validate_boolean v; @use_autowidth = v; end

    # inserts a worksheet into this workbook at the position specified.
    # It the index specified is out of range, the worksheet will be added to the end of the
    # worksheets collection
    # @return [Worksheet]
    # @param index The zero based position to insert the newly created worksheet
    # @param [Hash] options Options to pass into the worksheed during initialization.
    # @option options [String] name The name of the worksheet
    # @option options [Hash] page_margins The page margins for the worksheet
    def insert_worksheet(index=0, options={})
      worksheet = Worksheet.new(self, options)
      @worksheets.delete_at(@worksheets.size - 1)
      @worksheets.insert(index, worksheet)
      yield worksheet if block_given?
      worksheet
    end

    #
    # Adds a worksheet to this workbook
    # @return [Worksheet]
    # @option options [String] name The name of the worksheet.
    # @option options [Hash] page_margins The page margins for the worksheet.
    # @see Worksheet#initialize
    def add_worksheet(options={})
      worksheet = Worksheet.new(self, options)
      yield worksheet if block_given?
      worksheet
    end

    # Adds a defined name to this workbook
    # @return [DefinedName]
    # @param [String] formula @see DefinedName
    # @param [Hash] options @see DefinedName
    def add_defined_name(formula, options)
      defined_names << DefinedName.new(formula, options)
    end

    # The workbook relationships. This is managed automatically by the workbook
    # @return [Relationships]
    def relationships
      r = Relationships.new
      @worksheets.each do |sheet|
        r << Relationship.new(sheet, WORKSHEET_R, WORKSHEET_PN % (r.size+1))
      end
      pivot_tables.each_with_index do |pivot_table, index|
        r << Relationship.new(pivot_table.cache_definition, PIVOT_TABLE_CACHE_DEFINITION_R, PIVOT_TABLE_CACHE_DEFINITION_PN % (index+1))
      end
      r << Relationship.new(self, STYLES_R,  STYLES_PN)
      if use_shared_strings
          r << Relationship.new(self, SHARED_STRINGS_R,  SHARED_STRINGS_PN)
      end
      r
    end

    # generates a shared string object against all cells in all worksheets.
    # @return [SharedStringTable]
    def shared_strings
      SharedStringsTable.new(worksheets.collect { |ws| ws.cells }, xml_space)
    end

    # The xml:space attribute for the worksheet.
    # This determines how whitespace is handled withing the document.
    # The most relevant part being whitespace in the cell text.
    # allowed values are :preserve and :default. Axlsx uses :preserve unless
    # you explicily set this to :default.
    # @return Symbol
    def xml_space
      @xml_space ||= :preserve
    end

    # Sets the xml:space attribute for the worksheet
    # @see Worksheet#xml_space
    # @param [Symbol] space must be one of :preserve or :default
    def xml_space=(space)
      Axlsx::RestrictionValidator.validate(:xml_space, [:preserve, :default], space)
      @xml_space = space;
    end

    # returns a range of cells in a worksheet
    # @param [String] cell_def The excel style reference defining the worksheet and cells. The range must specify the sheet to
    # retrieve the cells from. e.g. range('Sheet1!A1:B2') will return an array of four cells [A1, A2, B1, B2] while range('Sheet1!A1') will return a single Cell.
    # @return [Cell, Array]
    def [](cell_def)
      sheet_name = cell_def.split('!')[0] if cell_def.match('!')
      worksheet =  self.worksheets.select { |s| s.name == sheet_name }.first
      raise ArgumentError, 'Unknown Sheet' unless sheet_name && worksheet.is_a?(Worksheet)
      worksheet[cell_def.gsub(/.+!/,"")]
    end

    # Serialize the workbook
    # @param [String] str
    # @return [String]
    def to_xml_string(str='')
      add_worksheet unless worksheets.size > 0
      str << '<?xml version="1.0" encoding="UTF-8"?>'
      str << '<workbook xmlns="' << XML_NS << '" xmlns:r="' << XML_NS_R << '">'
      str << '<workbookPr date1904="' << @@date1904.to_s << '"/>'
      str << '<sheets>'
      @worksheets.each_with_index do |sheet, index|
        str << '<sheet name="' << sheet.name << '" sheetId="' << (index+1).to_s << '" r:id="' << sheet.rId << '"/>'
        if defined_name = sheet.auto_filter.defined_name
          add_defined_name defined_name, :name => '_xlnm._FilterDatabase', :local_sheet_id => index, :hidden => 1
        end
      end
      str << '</sheets>'
      defined_names.to_xml_string(str)
      unless pivot_tables.empty?
        str << '<pivotCaches>'
        pivot_tables.each do |pivot_table|
          str << '<pivotCache cacheId="' << pivot_table.cache_definition.cache_id.to_s << '" r:id="' << pivot_table.cache_definition.rId << '"/>'
        end
        str << '</pivotCaches>'
      end
      str << '</workbook>'
    end

  end
end
