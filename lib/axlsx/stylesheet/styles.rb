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
    # @option options [Boolean] strike Indicates if the text should be rendered with a strikethrough
    # @option options [Boolean] strike Indicates if the text should be rendered with a shadow
    # @option options [Integer] charset The character set to use.
    # @option options [Integer] family The font family to use.
    # @option options [String] name The name of the font to use
    # @option options [Integer] num_fmt The number format to apply
    # @option options [String] format_code The formatting to apply. If this is specified, num_fmt is ignored.
    # @option options [Integer] border The border style to use. 
    # @option options [String] bg_color The background color to apply to the cell
    # @option options [Boolean] hidden Indicates if the cell should be hidden
    # @option options [Boolean] locked Indicates if the cell should be locked
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
    #   title = ws.style.add_style(:bg_color => "FFFF0000", :fg_color=>"#FF000000", :sz=>14,  :border=>Axlsx::STYLE_THIN_BORDER
    #
    #   ws.add_row :values => ["Least Popular Pets"]
    #   ws.add_row :values => ["", "Dry Skinned Reptiles", "Bald Cats", "Violent Parrots"], :style=>title
    #   ws.add_row :values => ["Votes", 6, 4, 1], :style=>Axlsx::STYLE_THIN_BORDER
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
    #   title = ws.style.add_style(:bg_color => "FFFF0000", 
    #                              :fg_color=>"#FF000000",
    #                              :border=>Axlsx::STYLE_THIN_BORDER, 
    #                              :alignment=>{:horizontal => :center})
    #
    #   date_time = ws.style.add_style(:num_fmt => Axlsx::NUM_FMT_YYYYMMDDHHMMSS,
    #                                  :border=>Axlsx::STYLE_THIN_BORDER)
    #
    #   percent = ws.style.add_style(:num_fmt => Axlsx::NUM_FMT_PERCENT, 
    #                                :border=>Axlsx::STYLE_THIN_BORDER)
    #
    #   currency = ws.style.add_style(:format_code=>"¥#,##0;[Red]¥-#,##0",
    #                                 :border=>Axlsx::STYLE_THIN_BORDER)
    #
    #   # build your rows
    #   ws.add_row :values => ["Genreated At:", Time.now], :styles=>[nil, date_time]
    #   ws.add_row :values => ["Previous Year Quarterly Profits (JPY)"], :style=>title
    #   ws.add_row :values => ["Quarter", "Profit", "% of Total"], :style=>title
    #   ws.add_row :values => ["Q1", 4000, 40], :style=>[title, currency, percent]
    #   ws.add_row :values => ["Q2", 3000, 30], :style=>[title, currency, percent]
    #   ws.add_row :values => ["Q3", 1000, 10], :style=>[title, currency, percent]
    #   ws.add_row :values => ["Q4", 2000, 20], :style=>[title, currency, percent]
    #   f = File.open('example_you_got_style.xlsx', 'w')
    #   p.serialize(f)
    def add_style(options={})
      
      numFmtId = if options[:format_code]
                   n = @numFmts.map{ |f| f.numFmtId }.max + 1
                   numFmts << NumFmt.new(:numFmtId => n, :formatCode=> options[:format_code])
                   n
                 else
                   options[:num_fmt] || 0
                 end
      
      borderId = options[:border] || 0
      raise ArgumentError, "Invalid borderId" unless borderId < borders.size
      
      fill = if options[:bg_color]
               color = Color.new(:rgb=>options[:bg_color])
               pattern = PatternFill.new(:patternType =>:solid, :fgColor=>color)
               fills << Fill.new(pattern)   
             else
               0
             end
      
      fontId = if (options.values_at(:fg_color, :sz, :b, :i, :strike, :outline, :shadow, :charset, :family, :font_name).length)
                 font = Font.new()
                 [:b, :i, :strike, :outline, :shadow, :charset, :family, :sz].each { |k| font.send("#{k}=", options[k]) unless options[k].nil? }
                 font.color = Color.new(:rgb => options[:fg_color]) unless options[:fg_color].nil?
                 font.name = options[:font_name] unless options[:font_name].nil?
                 fonts << font
               else
                 0 
               end
      
      applyProtection = (options[:hidden] || options[:locked]) ? 1 : 0
      
      xf = Xf.new(:fillId => fill, :fontId=>fontId, :applyFill=>1, :applyFont=>1, :numFmtId=>numFmtId, :borderId=>borderId, :applyProtection=>applyProtection)

      xf.applyNumberFormat = true if xf.numFmtId > 0
      
      if options[:alignment]
        xf.alignment = CellAlignment.new(options[:alignment])
      end
      
      if applyProtection
        xf.protection = CellProtection.new(options)
      end
      
      cellXfs << xf
    end
    
    # Serializes the styles document
    # @return [String]
    def to_xml()
      builder = Nokogiri::XML::Builder.new(:encoding => ENCODING) do |xml|
        xml.styleSheet(:xmlns => XML_NS) {
          [:numFmts, :fonts, :fills, :borders, :cellStyleXfs, :cellXfs, :cellStyles, :dxfs, :tableStyles].each do |key|
            self.instance_values[key.to_s].to_xml(xml) unless self.instance_values[key.to_s].nil?
          end
        }
      end
      builder.to_xml(:save_with => 0)
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

      @dxfs = SimpleTypedList.new(Xf, "dxfs"); @dxfs.lock
      @tableStyles = TableStyles.new(:defaultTableStyle => "TableStyleMedium9", :defaultPivotStyle => "PivotStyleLight16"); @tableStyles.lock
    end
  end
end
  
