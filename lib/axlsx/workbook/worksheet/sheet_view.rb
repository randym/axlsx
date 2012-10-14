# encoding: UTF-8
module Axlsx
  # View options for a worksheet.
  #
  # @note The recommended way to manage the sheet view is via Worksheet#sheet_view
  # @see Worksheet#sheet_view
  class SheetView
   
    include Axlsx::OptionsParser
    include Axlsx::Accessors
    include Axlsx::SerializedAttributes

    # Creates a new {SheetView} object
    # @option options [Integer] color_id Color Id
    # @option options [Boolean] default_grid_color Default Grid Color
    # @option options [Boolean] right_to_left Right To Left
    # @option options [Boolean] show_formulas Show Formulas
    # @option options [Boolean] show_grid_lines Show Grid Lines
    # @option options [Boolean] show_outline_symbols Show Outline Symbols
    # @option options [Boolean] show_row_col_headers Show Headers
    # @option options [Boolean] show_ruler Show Ruler
    # @option options [Boolean] show_white_space Show White Space
    # @option options [Boolean] show_zeros Show Zero Values
    # @option options [Boolean] tab_selected Sheet Tab Selected
    # @option options [String, Cell] top_left_cell Top Left Visible Cell
    # @option options [Symbol] view View Type
    # @option options [Boolean] window_protection Window Protection
    # @option options [Integer] workbook_view_id Workbook View Index
    # @option options [Integer] zoom_scale
    # @option options [Integer] zoom_scale_normal Zoom Scale Normal View
    # @option options [Integer] zoom_scale_page_layout_view Zoom Scale Page Layout View
    # @option options [Integer] zoom_scale_sheet_layout_view Zoom Scale Page Break Preview
    def initialize(options={})
      #defaults
      @color_id = @top_left_cell = @pane = nil
      @right_to_left = @show_formulas = @show_outline_symbols = @show_white_space = @tab_selected = @window_protection = false
      @default_grid_color = @show_grid_lines = @show_row_col_headers = @show_ruler = @show_zeros = true
      @zoom_scale = 100
      @zoom_scale_normal = @zoom_scale_page_layout_view = @zoom_scale_sheet_layout_view = @workbook_view_id = 0
      @selections = {}
      parse_options options
    end

    boolean_attr_accessor :default_grid_color, :right_to_left, :show_formulas, :show_grid_lines,
      :show_row_col_headers, :show_ruler, :show_white_space, :show_zeros, :tab_selected, :window_protection, :show_outline_symbols

    serializable_attributes :default_grid_color, :right_to_left, :show_formulas, :show_grid_lines,
      :show_row_col_headers, :show_ruler, :show_white_space, :show_zeros, :tab_selected, :window_protection, :show_outline_symbols,
      :zoom_scale_sheet_layout_view, :zoom_scale_page_layout_view, :zoom_scale_normal, :workbook_view_id,
      :view, :top_left_cell, :color_id, :zoom_scale


    # instance values that must be serialized as their own elements - e.g. not attributes.
    CHILD_ELEMENTS = [ :pane, :selections ]

    # The pane object for the sheet view
    # @return [Pane]
    # @see [Pane]
    def pane
      @pane ||= Pane.new
      yield @pane if block_given?
      @pane
    end

    # A hash of selection objects keyed by pane type associated with this sheet view.
    # @return [Hash]
    attr_reader :selections

    #  
    # Color Id
    # Index to the color value for row/column
    # text headings and gridlines. This is an 
    # 'index color value' (ICV) rather than 
    # rgb value.
    # @see type
    # @return [Integer]
    # default nil
    attr_reader :color_id

    # Top Left Visible Cell
    # Location of the top left visible cell Location 
    # of the top left visible cell in the bottom right
    # pane (when in Left-to-Right mode).
    # @see type
    # @return [String]
    # default nil
    attr_reader :top_left_cell
    
    
    # View Type
    # Indicates the view type.
    # Options are 
    #  * normal: Normal view
    #  * page_break_preview: Page break preview
    #  * page_layout: Page Layout View
    # @see type
    # @return [Symbol]
    # default :normal
    attr_reader :view

    # Workbook View Index
    # Zero-based index of this workbook view, pointing 
    # to a workbookView element in the bookViews collection.
    # @see type
    # @return [Integer] 
    # default 0
    attr_reader :workbook_view_id

    # Zoom Scale
    # Window zoom magnification for current view 
    # representing percent values. This attribute
    # is restricted to values ranging from 10 to 400. 
    # Horizontal & Vertical scale together.
    # Current view can be Normal, Page Layout, or 
    # Page Break Preview.
    # @see type
    # @return [Integer] 
    # default 100
    attr_reader :zoom_scale


    # Zoom Scale Normal View
    # Zoom magnification to use when in normal view, 
    # representing percent values. This attribute is 
    # restricted to values ranging from 10 to 400. 
    # Horizontal & Vertical scale together.
    # Applies for worksheets only; zero implies the 
    # automatic setting.
    # @see type
    # @return [Integer] 
    # default 0
    attr_reader :zoom_scale_normal


    # Zoom Scale Page Layout View
    # Zoom magnification to use when in page layout 
    # view, representing percent values. This attribute 
    # is restricted to values ranging from 10 to 400. 
    # Horizontal & Vertical scale together.
    # Applies for worksheets only; zero implies 
    # the automatic setting.
    # @see type
    # @return [Integer] 
    # default 0
    attr_reader :zoom_scale_page_layout_view


    # Zoom Scale Page Break Preview
    # Zoom magnification to use when in page break 
    # preview, representing percent values. This 
    # attribute is restricted to values ranging 
    # from 10 to 400. Horizontal & Vertical scale
    # together.
    # Applies for worksheet only; zero implies 
    # the automatic setting.
    # @see type
    # @return [Integer] 
    # default 0
    attr_reader :zoom_scale_sheet_layout_view

    # Adds a new selection
    # param [Symbol] pane
    # param [Hash] options
    # return [Selection]
    def add_selection(pane, options = {})
      @selections[pane] = Selection.new(options.merge(:pane => pane))
    end

    # @see color_id
    def color_id=(v); Axlsx::validate_unsigned_int(v); @color_id = v end

    # @see top_left_cell
    def top_left_cell=(v)
      cell = (v.class == Axlsx::Cell ? v.r_abs : v)
      Axlsx::validate_string(cell)  
      @top_left_cell = cell
    end

    # @see view
    def view=(v); Axlsx::validate_sheet_view_type(v); @view = v end

    # @see workbook_view_id
    def workbook_view_id=(v); Axlsx::validate_unsigned_int(v); @workbook_view_id = v end

    # @see zoom_scale
    def zoom_scale=(v); Axlsx::validate_scale_0_10_400(v); @zoom_scale = v end

    # @see zoom_scale_normal
    def zoom_scale_normal=(v); Axlsx::validate_scale_0_10_400(v); @zoom_scale_normal = v end

    # @see zoom_scale_page_layout_view
    def zoom_scale_page_layout_view=(v); Axlsx::validate_scale_0_10_400(v); @zoom_scale_page_layout_view = v end

    # @see zoom_scale_sheet_layout_view
    def zoom_scale_sheet_layout_view=(v); Axlsx::validate_scale_0_10_400(v); @zoom_scale_sheet_layout_view = v end

    # Serializes the data validation
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<sheetViews>'
      str << '<sheetView '
      serialized_attributes str
      str << '>'
      @pane.to_xml_string(str) if @pane
      @selections.each do |key, selection|
        selection.to_xml_string(str)
      end
      str << '</sheetView>'
      str << '</sheetViews>'
    end
  end
end
