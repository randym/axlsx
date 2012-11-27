$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'
require 'date'

p = Axlsx::Package.new
wb = p.workbook
wb.styles do |style|

  # Date/Time Styles
  #
  # The most important thing to remember about OOXML styles is that they are
  # exclusive. This means that each style must define all the components it
  # requires to render the cell the way you want. A good example of this is
  # changing the font size for a date. You cannot specify just the font size,
  # you must also specify the number format or format code so that renders
  # know how to display the serialized date float value 
  #
  # The parts that make up a custom styles are:
  #
  # fonts(Font), fills(Fill), borders(Border) and number formats(NumFmt).
  # Getting to know those classes will help you make the most out of custom
  # styling. However axlsx certainly does not expect you to create all those
  # objects manually.
  #
  # workbook.styles.add_style provides a helper method 'add_style' for defining
  # styles in one go. The docs for that method are definitely worth a read.
  # @see Style#add_style

  # When no style is applied to a cell, axlsx will automatically apply date/time
  # formatting to Date and Time objects for you. However, if you are defining
  # custom styles, you define all aspects of the style you want to apply.
  #
  # An aside on styling and auto-width. Auto-width calculations do not
  # currently take into account any style or formatting you have applied to the
  # data in your cells as it would require the creation of a rendering engine,
  # and frankly kill performance. If you are doing a lot of custom formatting,
  # you are going to be better served by specifying fixed column widths.
  #
  # Let's look at an example:
  #
  # A style that only applies a font size
  large_font = wb.styles.add_style :sz => 20

  # A style that applies both a font size and a predefined number format.
  # @see NumFmt
  predefined_format = wb.styles.add_style :sz => 20, :num_fmt => 14

  # A style that a applies a font size and a custom formatting code
  custom_format = wb.styles.add_style :sz => 20, :format_code => 'yyyy-mm-dd'

  # A style that overrides top and left border style
  override_border = wb.styles.add_style :border => { :style => :thin, :color =>"FAAC58", :edges => [:right, :top, :left] }, :border_top => { :style => :thick, :color => "01DF74" }, :border_left => { :color => "0101DF" }


  wb.add_worksheet do |sheet|

    # We then apply those styles positionally
    sheet.add_row [123, "123", Time.now], style: [nil, large_font, predefined_format]
    sheet.add_row [123, "123", Date.new(2012, 9, 14)], style: [large_font, nil, custom_format]
    sheet.add_row [123, "123", Date.new(2000, 9, 12)] # This uses the axlsx default format_code (14)
    sheet.add_row [123, "123", Time.now], style: [large_font, override_border, predefined_format]
  end

end
p.serialize 'styles.xlsx'

