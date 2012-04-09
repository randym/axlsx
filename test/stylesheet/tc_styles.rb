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
    s = @styles.add_style :border => {:style=>:thin, :color => "FFFF0000"}
    assert_equal(@styles.borders.size, border_count + 1)
    assert_equal(@styles.borders.last.prs.last.color.rgb, "FFFF0000")
    assert_raise(ArgumentError) { @styles.add_style :border => {:color => "FFFF0000"} }
    [:top,:bottom,:left,:right].each {|edge| assert(@styles.borders.last.edges.include?(edge))}

    s2 = @styles.add_style :border => {:style=>:thin, :color => "0000FFFF", :edges => [:top, :bottom]}
    assert_equal(@styles.borders.size, border_count + 2)
    assert_equal(@styles.borders.last.prs.last.color.rgb, "0000FFFF")
    assert(@styles.borders.last.edges.include?(:top))
    assert(@styles.borders.last.edges.include?(:bottom))
    assert(! @styles.borders.last.edges.include?(:left))
    assert(! @styles.borders.last.edges.include?(:right))
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

end
