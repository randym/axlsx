require 'tc_helper.rb'

class TestStyles < Test::Unit::TestCase
  def setup
    @styles = Axlsx::Styles.new
  end
  def teardown
  end

  def test_valid_document
    schema = Nokogiri::XML::Schema(File.open(Axlsx::SML_XSD))
    doc = Nokogiri::XML(@styles.to_xml_string)
    errors = []
    schema.validate(doc).each do |error|
      errors.push error
      puts error.message
    end
    assert(errors.size == 0)
  end
  def test_add_style_border_hash
    border_count = @styles.borders.size
    s = @styles.add_style :border => {:style => :thin, :color => "FFFF0000"}
    assert_equal(@styles.borders.size, border_count + 1)
    assert_equal(@styles.borders.last.prs.last.color.rgb, "FFFF0000")
    assert_raise(ArgumentError) { @styles.add_style :border => {:color => "FFFF0000"} }
    assert_equal @styles.borders.last.prs.size, 4
  end

  def test_add_style_border_edges
    s = @styles.add_style :border => { :style => :thin, :color => "0000FFFF", :edges => [:top, :bottom] }
    parts = @styles.borders.last.prs
    parts.each { |pr| assert_equal(pr.color.rgb, "0000FFFF", "Style is applied to #{pr.name} properly") }
    assert((parts.map { |pr| pr.name.to_s }.sort && ['bottom', 'top']).size == 2, "specify two edges, and you get two border prs")
  end

  def test_do_not_alter_options_in_add_style
    #This should test all options, but for now - just the bits that we know caused some pain
    options = { :border => { :style => :thin, :color =>"FF000000" } }
    @styles.add_style options
    assert_equal options[:border][:style], :thin, 'thin style is stil in option'
    assert_equal options[:border][:color], "FF000000", 'color is stil in option'
  end

  def test_add_style
    fill_count = @styles.fills.size
    font_count = @styles.fonts.size
    xf_count = @styles.cellXfs.size

    @styles.add_style :bg_color=>"FF000000", :fg_color=>"FFFFFFFF", :sz=>13, :num_fmt=>Axlsx::NUM_FMT_PERCENT, :alignment=>{:horizontal=>:left}, :border=>Axlsx::STYLE_THIN_BORDER, :hidden=>true, :locked=>true
    assert_equal(@styles.fills.size, fill_count+1)
    assert_equal(@styles.fonts.size, font_count+1)
    assert_equal(@styles.cellXfs.size, xf_count+1)
    xf = @styles.cellXfs.last
    assert_equal(xf.fillId, (@styles.fills.size-1), "points to the last created fill")
    assert_equal(@styles.fills.last.fill_type.fgColor.rgb, "FF000000", "fill created with color")

    assert_equal(xf.fontId, (@styles.fonts.size-1), "points to the last created font")
    assert_equal(@styles.fonts.last.sz, 13, "font sz applied")
    assert_equal(@styles.fonts.last.color.rgb, "FFFFFFFF", "font color applied")

    assert_equal(xf.borderId, Axlsx::STYLE_THIN_BORDER, "border id is set")
    assert_equal(xf.numFmtId, Axlsx::NUM_FMT_PERCENT, "number format id is set")

    assert(xf.alignment.is_a?(Axlsx::CellAlignment), "alignment was created")
    assert_equal(xf.alignment.horizontal, :left, "horizontal alignment applied")
    assert_equal(xf.protection.hidden, true, "hidden protection set")
    assert_equal(xf.protection.locked, true, "cell locking set")
    assert_raise(ArgumentError, "should reject invalid borderId") { @styles.add_style :border => 2 }


    assert_equal(xf.applyProtection, 1, "protection applied")
    assert_equal(xf.applyBorder, true, "border applied")
    assert_equal(xf.applyNumberFormat, true, "number format applied")
    assert_equal(xf.applyAlignment, true, "alignment applied")
  end
  
  def test_basic_add_style_dxf
    border_count = @styles.borders.size
    s = @styles.add_style :border => {:style => :thin, :color => "FFFF0000"}, :type => :dxf
    assert_equal(@styles.borders.size, border_count, "styles borders not affected")
    assert_equal(@styles.dxfs.last.border.prs.last.color.rgb, "FFFF0000")
    assert_raise(ArgumentError) { @styles.add_style :border => {:color => "FFFF0000"}, :type => :dxf }
    assert_equal @styles.borders.last.prs.size, 4
  end

  def test_add_style_dxf
    fill_count = @styles.fills.size
    font_count = @styles.fonts.size
    dxf_count = @styles.dxfs.size

    style = @styles.add_style :bg_color=>"FF000000", :fg_color=>"FFFFFFFF", :sz=>13, :alignment=>{:horizontal=>:left}, :border=>{:style => :thin, :color => "FFFF0000"}, :hidden=>true, :locked=>true, :type => :dxf
    assert_equal(@styles.dxfs.size, dxf_count+1)
    assert_equal(0, style, "returns the zero-based dxfId")
    
    dxf = @styles.dxfs.last
    assert_equal(@styles.dxfs.last.fill.fill_type.fgColor.rgb, "FF000000", "fill created with color")

    assert_equal(font_count, (@styles.fonts.size), "font not created under styles")
    assert_equal(fill_count, (@styles.fills.size), "fill not created under styles")

    assert(dxf.border.is_a?(Axlsx::Border), "border is set")
    assert_equal(nil, dxf.numFmt, "number format is not set")

    assert(dxf.alignment.is_a?(Axlsx::CellAlignment), "alignment was created")
    assert_equal(dxf.alignment.horizontal, :left, "horizontal alignment applied")
    assert_equal(dxf.protection.hidden, true, "hidden protection set")
    assert_equal(dxf.protection.locked, true, "cell locking set")
    assert_raise(ArgumentError, "should reject invalid borderId") { @styles.add_style :border => 3 }
    assert_raise(ArgumentError, "should reject num_fmt option") {
      @styles.add_style :type=>:dxf, :num_fmt=>Axlsx::NUM_FMT_PERCENT
    }
  end

  def test_multiple_dxf
    # add a second style
    style = @styles.add_style :bg_color=>"00000000", :fg_color=>"FFFFFFFF", :sz=>13, :alignment=>{:horizontal=>:left}, :border=>{:style => :thin, :color => "FFFF0000"}, :hidden=>true, :locked=>true, :type => :dxf
    assert_equal(0, style, "returns the first dxfId")
    style = @styles.add_style :bg_color=>"FF000000", :fg_color=>"FFFFFFFF", :sz=>13, :alignment=>{:horizontal=>:left}, :border=>{:style => :thin, :color => "FFFF0000"}, :hidden=>true, :locked=>true, :type => :dxf
    assert_equal(1, style, "returns the second dxfId")
  end
end
