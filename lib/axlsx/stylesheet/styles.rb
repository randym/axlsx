# encoding: UTF-8
module Axlsx
  require 'axlsx/stylesheet/border.rb'
  require 'axlsx/stylesheet/border_pr.rb'
  require 'axlsx/stylesheet/cell_alignment.rb'
  require 'axlsx/stylesheet/cell_style.rb'
  require 'axlsx/stylesheet/color.rb'
  require 'axlsx/stylesheet/fill.rb'
  require 'axlsx/stylesheet/font.rb'
  require 'axlsx/stylesheet/gradient_fill.rb'
  require 'axlsx/stylesheet/gradient_stop.rb'
  require 'axlsx/stylesheet/num_fmt.rb'
  require 'axlsx/stylesheet/pattern_fill.rb'
  require 'axlsx/stylesheet/table_style.rb'
  require 'axlsx/stylesheet/table_styles.rb'
  require 'axlsx/stylesheet/table_style_element.rb'
  require 'axlsx/stylesheet/dxf.rb'
  require 'axlsx/stylesheet/xf.rb'
  require 'axlsx/stylesheet/cell_protection.rb'

  #The Styles class manages worksheet styles
  # In addition to creating the require style objects for a valid xlsx package, this class provides the key mechanism for adding styles to your workbook, and safely applying them to the cells of your worksheet.
  # All portions of the stylesheet are implemented here exception colors, which specify legacy and modified pallete colors, and exLst, whic is used as a future feature data storage area.
  # @see  Office Open XML Part 1 18.8.11 for gory details on how this stuff gets put together
  # @see  Styles#add_style
  # @note The recommended way to manage styles is with add_style
  class Styles
    # numFmts for your styles.
    #  The default styles, which change based on the system local, are as follows.
    #  id formatCode
    #   0 General
    #   1 0
    #   2 0.00
    #   3 #,##0
    #   4 #,##0.00
    #   9 0%
    #   10 0.00%
    #   11 0.00E+00
    #   12 #   ?/?
    #   13 #   ??/??
    #   14 mm-dd-yy
    #   15 d-mmm-yy
    #   16 d-mmm
    #   17 mmm-yy
    #   18 h:mm AM/PM
    #   19 h:mm:ss AM/PM
    #   20 h:mm
    #   21 h:mm:ss
    #   22 m/d/yy h:mm
    #   37 #,##0 ;(#,##0)
    #   38 #,##0 ;[Red](#,##0)
    #   39 #,##0.00;(#,##0.00)
    #   40 #,##0.00;[Red](#,##0.00)
    #   45 mm:ss
    #   46 [h]:mm:ss
    #   47 mmss.0
    #   48 ##0.0E+0
    #   49 @
    #  Axlsx also defines the following constants which you can use in add_style.
    #     NUM_FMT_PERCENT formats to "0%"
    #     NUM_FMT_YYYYMMDD formats to "yyyy/mm/dd"
    #     NUM_FMT_YYYYMMDDHHMMSS  formats to "yyyy/mm/dd hh:mm:ss"
    # @see Office Open XML Part 1 - 18.8.31 for more information on creating number formats
    # @return [SimpleTypedList]
    # @note The recommended way to manage styles is with add_style
    # @see Styles#add_style
    attr_reader :numFmts

    # The collection of fonts used in this workbook
    # @return [SimpleTypedList]
    # @note The recommended way to manage styles is with add_style
    # @see Styles#add_style
    attr_reader :fonts

    # The collection of fills used in this workbook
    # @return [SimpleTypedList]
    # @note The recommended way to manage styles is with add_style
    # @see Styles#add_style
    attr_reader :fills

    # The collection of borders used in this workbook
    # Axlsx predefines THIN_BORDER which can be used to put a border around all of your cells.
    # @return [SimpleTypedList]
    # @note The recommended way to manage styles is with add_style
    # @see Styles#add_style
    attr_reader :borders

    # The collection of master formatting records for named cell styles, which means records defined in cellStyles, in the workbook
    # @return [SimpleTypedList]
    # @note The recommended way to manage styles is with add_style
    # @see Styles#add_style
    attr_reader :cellStyleXfs

    # The collection of named styles, referencing cellStyleXfs items in the workbook.
    # @return [SimpleTypedList]
    # @note The recommended way to manage styles is with add_style
    # @see  Styles#add_style
    attr_reader :cellStyles

    # The collection of master formatting records. This is the list that you will actually use in styling a workbook.
    # @return [SimpleTypedList]
    # @note The recommended way to manage styles is with add_style
    # @see Styles#add_style
    attr_reader :cellXfs

    # The collection of non-cell formatting records used in the worksheet.
    # @return [SimpleTypedList]
    # @note The recommended way to manage styles is with add_style
    # @see Styles#add_style
    attr_reader :dxfs

    # The collection of table styles that will be available to the user in the excel UI
    # @return [SimpleTypedList]
    # @note The recommended way to manage styles is with add_style
    # @see Styles#add_style
    attr_reader :tableStyles

    # Creates a new Styles object and prepopulates it with the requires objects to generate a valid package style part.
    def initialize()
      load_default_styles
    end

    # Drastically simplifies style creation and management.
    # @return [Integer]
    # @option options [String] fg_color The text color
    # @option options [Integer] sz The text size
    # @option options [Boolean] b Indicates if the text should be bold
    # @option options [Boolean] i Indicates if the text should be italicised
    # @option options [Boolean] u Indicates if the text should be underlined
    # @option options [Boolean] strike Indicates if the text should be rendered with a strikethrough
    # @option options [Boolean] shadow Indicates if the text should be rendered with a shadow
    # @option options [Integer] charset The character set to use.
    # @option options [Integer] family The font family to use.
    # @option options [String] font_name The name of the font to use
    # @option options [Integer] num_fmt The number format to apply
    # @option options [String] format_code The formatting to apply. 
    # @option options [Integer|Hash] border The border style to use.
    #   borders support style, color and edges options @see parse_border_options
    # @option options [String] bg_color The background color to apply to the cell
    # @option options [Boolean] hidden Indicates if the cell should be hidden
    # @option options [Boolean] locked Indicates if the cell should be locked
    # @option options [Symbol] type What type of style is this. Options are [:dxf, :xf]. :xf is default
    # @option options [Hash] alignment A hash defining any of the attributes used in CellAlignment
    # @see CellAlignment
    #
    # @example You Got Style
    #   require "rubygems" # if that is your preferred way to manage gems!
    #   require "axlsx"
    #
    #   p = Axlsx::Package.new
    #   ws = p.workbook.add_worksheet
    #
    #   # black text on a white background at 14pt with thin borders!
    #   title = ws.styles.add_style(:bg_color => "FFFF0000", :fg_color=>"#FF000000", :sz=>14,  :border=> {:style => :thin, :color => "FFFF0000"}
    #
    #   ws.add_row ["Least Popular Pets"]
    #   ws.add_row ["", "Dry Skinned Reptiles", "Bald Cats", "Violent Parrots"], :style=>title
    #   ws.add_row ["Votes", 6, 4, 1], :style=>Axlsx::STYLE_THIN_BORDER
    #   f = File.open('example_you_got_style.xlsx', 'w')
    #   p.serialize(f)
    #
    # @example Styling specifically
    #   # an example of applying specific styles to specific cells
    #   require "rubygems" # if that is your preferred way to manage gems!
    #   require "axlsx"
    #
    #   p = Axlsx::Package.new
    #   ws = p.workbook.add_worksheet
    #
    #   # define your styles
    #   title = ws.styles.add_style(:bg_color => "FFFF0000",
    #                              :fg_color=>"#FF000000",
    #                              :border=>Axlsx::STYLE_THIN_BORDER,
    #                              :alignment=>{:horizontal => :center})
    #
    #   date_time = ws.styles.add_style(:num_fmt => Axlsx::NUM_FMT_YYYYMMDDHHMMSS,
    #                                  :border=>Axlsx::STYLE_THIN_BORDER)
    #
    #   percent = ws.styles.add_style(:num_fmt => Axlsx::NUM_FMT_PERCENT,
    #                                :border=>Axlsx::STYLE_THIN_BORDER)
    #
    #   currency = ws.styles.add_style(:format_code=>"¥#,##0;[Red]¥-#,##0",
    #                                 :border=>Axlsx::STYLE_THIN_BORDER)
    #
    #   # build your rows
    #   ws.add_row ["Generated At:", Time.now], :styles=>[nil, date_time]
    #   ws.add_row ["Previous Year Quarterly Profits (JPY)"], :style=>title
    #   ws.add_row ["Quarter", "Profit", "% of Total"], :style=>title
    #   ws.add_row ["Q1", 4000, 40], :style=>[title, currency, percent]
    #   ws.add_row ["Q2", 3000, 30], :style=>[title, currency, percent]
    #   ws.add_row ["Q3", 1000, 10], :style=>[title, currency, percent]
    #   ws.add_row ["Q4", 2000, 20], :style=>[title, currency, percent]
    #   f = File.open('example_you_got_style.xlsx', 'w')
    #   p.serialize(f)
    #
    # @example Differential styling
    #   # Differential styles apply on top of cell styles. Used in Conditional Formatting. Must specify :type => :dxf, and you can't use :num_fmt.
    #   require "rubygems" # if that is your preferred way to manage gems!
    #   require "axlsx"
    #
    #   p = Axlsx::Package.new
    #   wb = p.workbook
    #   ws = wb.add_worksheet
    #
    #   # define your styles
    #   profitable = wb.styles.add_style(:bg_color => "FFFF0000",
    #                              :fg_color=>"#FF000000",
    #                              :type => :dxf)
    #
    #   ws.add_row ["Genreated At:", Time.now], :styles=>[nil, date_time]
    #   ws.add_row ["Previous Year Quarterly Profits (JPY)"], :style=>title
    #   ws.add_row ["Quarter", "Profit", "% of Total"], :style=>title
    #   ws.add_row ["Q1", 4000, 40], :style=>[title, currency, percent]
    #   ws.add_row ["Q2", 3000, 30], :style=>[title, currency, percent]
    #   ws.add_row ["Q3", 1000, 10], :style=>[title, currency, percent]
    #   ws.add_row ["Q4", 2000, 20], :style=>[title, currency, percent]
    #
    #   ws.add_conditional_formatting("A1:A7", { :type => :cellIs, :operator => :greaterThan, :formula => "2000", :dxfId => profitable, :priority => 1 })
    #   f = File.open('example_differential_styling', 'w')
    #   p.serialize(f)
    #
    def add_style(options={})
      # Default to :xf
      options[:type] ||= :xf
      raise ArgumentError, "Type must be one of [:xf, :dxf]" unless [:xf, :dxf].include?(options[:type] )

      fill = parse_fill_options options
      font = parse_font_options options
      numFmt = parse_num_fmt_options options
      border = parse_border_options options
      alignment = parse_alignment_options options
      protection = parse_protection_options options

      case options[:type]
      when :dxf
        style = Dxf.new :fill => fill, :font => font, :numFmt => numFmt, :border => border, :alignment => alignment, :protection => protection
      else
        style = Xf.new :fillId=>fill || 0, :fontId=>font || 0, :numFmtId=>numFmt || 0, :borderId=>border || 0, :alignment => alignment, :protection => protection, :applyFill=>!fill.nil?, :applyFont=>!font.nil?, :applyNumberFormat =>!numFmt.nil?, :applyBorder=>!border.nil?, :applyAlignment => !alignment.nil?, :applyProtection => !protection.nil?
      end

      options[:type] == :xf ? cellXfs << style : dxfs << style
    end

    # parses add_style options for protection styles
    # noop if options hash does not include :hide or :locked key

    # @option options [Boolean] hide boolean value defining cell protection attribute for hiding.
    # @option options [Boolean] locked boolean value defining cell protection attribute for locking.
    # @return [CellProtection]
    def parse_protection_options(options={})
      return if (options.keys & [:hidden, :locked]).empty?
      CellProtection.new(options)
    end

    # parses add_style options for alignment
    # noop if options hash does not include :alignment key
    # @option options [Hash] alignment A hash of options to prive the CellAlignment intializer
    # @return [CellAlignment]
    # @see CellAlignment
    def parse_alignment_options(options={})
      return unless options[:alignment]
      CellAlignment.new options[:alignment]
    end

    # parses add_style options for fonts. If the options hash contains :type => :dxf we return a new Font object.
    # if not, we return the index of the newly created font object in the styles.fonts collection.
    # @note noop if none of the options described here are set on the options parameter.
    # @option options [Symbol] type The type of style object we are working with (dxf or xf)
    # @option options [String] fg_color The text color
    # @option options [Integer] sz The text size
    # @option options [Boolean] b Indicates if the text should be bold
    # @option options [Boolean] i Indicates if the text should be italicised
    # @option options [Boolean] u Indicates if the text should be underlined
    # @option options [Boolean] strike Indicates if the text should be rendered with a strikethrough
    # @option options [Boolean] outline Indicates if the text should be rendered with a shadow
    # @option options [Integer] charset The character set to use.
    # @option options [Integer] family The font family to use.
    # @option options [String] font_name The name of the font to use
    # @return [Font|Integer]
    def parse_font_options(options={})
      return if (options.keys & [:fg_color, :sz, :b, :i, :u, :strike, :outline, :shadow, :charset, :family, :font_name]).empty?
      fonts.first.instance_values.each do |key, value|
        # Thanks for that 1.8.7 - cant do a simple merge...
        options[key.to_sym] = value unless options.keys.include?(key.to_sym)
      end
      font = Font.new(options)
      font.color = Color.new(:rgb => options[:fg_color]) if options[:fg_color]
      font.name = options[:font_name] if options[:font_name]
      options[:type] == :dxf ? font : fonts << font
    end

    # parses add_style options for fills. If the options hash contains :type => :dxf we return a Fill object. If not, we return the index of the fill after being added to the fills collection.
    # @note noop if :bg_color is not specified in options
    # @option options [String] bg_color The rgb color to apply to the fill
    # @return [Fill|Integer]
    def parse_fill_options(options={})
      return unless options[:bg_color]
      color = Color.new(:rgb=>options[:bg_color])
      dxf = options[:type] == :dxf
      color_key = dxf ? :bgColor : :fgColor
      pattern = PatternFill.new(:patternType =>:solid, color_key=>color)
      fill = Fill.new(pattern)
      dxf ? fill : fills << fill
    end

    # parses Style#add_style options for borders.
    # @note noop if :border is not specified in options
    # @option options [Hash|Integer] A border style definition hash or the index of an existing border. 
    # Border style definition hashes must include :style and :color key-value entries and 
    # may include an :edges entry that references an array of symbols identifying which border edges 
    # you wish to apply the style or any other valid Border initializer options.
    # If the :edges entity is not provided the style is applied to all edges of cells that reference this style.
	# Also available :border_top, :border_right, :border_bottom and :border_left options with :style and/or :color 
	# key-value entries, which override :border values.
    # @example
    #   #apply a thick red border to the top and bottom
    #   { :border => { :style => :thick, :color => "FFFF0000", :edges => [:top, :bottom] }
    # @return [Border|Integer]
    def parse_border_options(options={})
      return unless options[:border]
      b_opts = options[:border]
      if b_opts.is_a?(Hash)
        raise ArgumentError, (ERR_INVALID_BORDER_OPTIONS % b_opts) unless b_opts.keys.include?(:style) && b_opts.keys.include?(:color)
        border = Border.new b_opts
        (b_opts[:edges] || [:left, :right, :top, :bottom]).each do |edge|
          edge_options = options["border_#{edge}".to_sym] || {}
          border_edge = b_opts.merge(edge_options)
          b_options = { :name => edge, :style => border_edge[:style], :color => Color.new(:rgb => border_edge[:color]) }
          border.prs << BorderPr.new(b_options)
        end
        options[:type] == :dxf ? border : borders << border
      elsif b_opts.is_a? Integer
        raise ArgumentError, (ERR_INVALID_BORDER_ID % b_opts) unless b_opts < borders.size
        if options[:type] == :dxf
          borders[b_opts].clone
        else
          border = b_opts
        end
      end
    end

    # Parses Style#add_style options for number formatting.
    # noop if neither :format_code or :num_format options are set.
    # @option options [Hash] A hash describing the :format_code and/or :num_fmt integer for the style.
    # @return [NumFmt|Integer]
    def parse_num_fmt_options(options={})
      return if (options.keys & [:format_code, :num_fmt]).empty?

      #When the user provides format_code - we always need to create a new numFmt object
      #When the type is :dxf we always need to create a new numFmt object
      if options[:format_code] || options[:type] == :dxf
        #If this is a standard xf we pull from numFmts the highest current and increment for num_fmt
        options[:num_fmt] ||= (@numFmts.map{ |num_fmt| num_fmt.numFmtId }.max + 1) if options[:type] != :dxf
        numFmt = NumFmt.new(:numFmtId => options[:num_fmt] || 0, :formatCode=> options[:format_code].to_s)
        options[:type] == :dxf ? numFmt : (numFmts << numFmt; numFmt.numFmtId)
      else
        options[:num_fmt]
      end
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << ('<styleSheet xmlns="' << XML_NS << '">')
      [:numFmts, :fonts, :fills, :borders, :cellStyleXfs, :cellXfs, :cellStyles, :dxfs, :tableStyles].each do |key|
        self.instance_values[key.to_s].to_xml_string(str) unless self.instance_values[key.to_s].nil?
      end
      str << '</styleSheet>'
    end

    private
    # Creates the default set of styles the exel requires to be valid as well as setting up the
    # Axlsx::STYLE_THIN_BORDER
    def load_default_styles
      @numFmts = SimpleTypedList.new NumFmt, 'numFmts'
      @numFmts << NumFmt.new(:numFmtId => NUM_FMT_YYYYMMDD, :formatCode=> "yyyy/mm/dd")
      @numFmts << NumFmt.new(:numFmtId => NUM_FMT_YYYYMMDDHHMMSS, :formatCode=> "yyyy/mm/dd hh:mm:ss")

      @numFmts.lock

      @fonts = SimpleTypedList.new Font, 'fonts'
      @fonts << Font.new(:name => "Arial", :sz => 11, :family=>1)
      @fonts.lock

      @fills = SimpleTypedList.new Fill, 'fills'
      @fills << Fill.new(Axlsx::PatternFill.new(:patternType=>:none))
      @fills << Fill.new(Axlsx::PatternFill.new(:patternType=>:gray125))
      @fills.lock

      @borders = SimpleTypedList.new Border, 'borders'
      @borders << Border.new
      black_border = Border.new
      [:left, :right, :top, :bottom].each do |item|
        black_border.prs << BorderPr.new(:name=>item, :style=>:thin, :color=>Color.new(:rgb=>"FF000000"))
      end
      @borders << black_border
      @borders.lock

      @cellStyleXfs = SimpleTypedList.new Xf, "cellStyleXfs"
      @cellStyleXfs << Xf.new(:borderId=>0, :numFmtId=>0, :fontId=>0, :fillId=>0)
      @cellStyleXfs.lock

      @cellStyles = SimpleTypedList.new CellStyle, 'cellStyles'
      @cellStyles << CellStyle.new(:name =>"Normal", :builtinId =>0, :xfId=>0)
      @cellStyles.lock

      @cellXfs = SimpleTypedList.new Xf, "cellXfs"
      @cellXfs << Xf.new(:borderId=>0, :xfId=>0, :numFmtId=>0, :fontId=>0, :fillId=>0)
      @cellXfs << Xf.new(:borderId=>1, :xfId=>0, :numFmtId=>0, :fontId=>0, :fillId=>0)
      # default date formatting
      @cellXfs << Xf.new(:borderId=>0, :xfId=>0, :numFmtId=>14, :fontId=>0, :fillId=>0, :applyNumberFormat=>1)
      @cellXfs.lock

      @dxfs = SimpleTypedList.new(Dxf, "dxfs"); @dxfs.lock
      @tableStyles = TableStyles.new(:defaultTableStyle => "TableStyleMedium9", :defaultPivotStyle => "PivotStyleLight16"); @tableStyles.lock
    end
  end
end

