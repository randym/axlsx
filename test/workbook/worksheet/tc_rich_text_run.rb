require 'tc_helper.rb'

class RichTextRun < Test::Unit::TestCase
  def setup
    @p = Axlsx::Package.new
    @ws = @p.workbook.add_worksheet :name => "hmmmz"
    @p.workbook.styles.add_style :sz => 20
    @rtr = Axlsx::RichTextRun.new('hihihi', b: true, i: false)
    @rtr2 = Axlsx::RichTextRun.new('hihi2hi2', b: false, i: true)
    @rt = Axlsx::RichText.new 
    @rt.runs << @rtr
    @rt.runs << @rtr2
    @row = @ws.add_row [@rt]
    @c = @row.first
  end

  def test_initialize
    assert_equal(@rtr.value, 'hihihi')
    assert_equal(@rtr.b, true)
    assert_equal(@rtr.i, false)
  end
  
  def test_font_size_with_custom_style_and_no_sz
    @c.style = @c.row.worksheet.workbook.styles.add_style :bg_color => 'FF00FF'
    sz = @rtr.send(:font_size)
    assert_equal(sz, @c.row.worksheet.workbook.styles.fonts.first.sz * 1.5)
    sz = @rtr2.send(:font_size)
    assert_equal(sz, @c.row.worksheet.workbook.styles.fonts.first.sz)
  end

  def test_font_size_with_bolding
    @c.style = @c.row.worksheet.workbook.styles.add_style :b => true
    assert_equal(@c.row.worksheet.workbook.styles.fonts.first.sz * 1.5, @rtr.send(:font_size))
    assert_equal(@c.row.worksheet.workbook.styles.fonts.first.sz * 1.5, @rtr2.send(:font_size)) # is this the correct behaviour?
  end

  def test_font_size_with_custom_sz
    @c.style = @c.row.worksheet.workbook.styles.add_style :sz => 52
    sz = @rtr.send(:font_size)
    assert_equal(sz, 52 * 1.5)
    sz2 = @rtr2.send(:font_size)
    assert_equal(sz2, 52)
  end
  
  def test_rtr_with_sz
    @rtr.sz = 25
    assert_equal(25, @rtr.send(:font_size))
  end
  
  def test_color
    assert_raise(ArgumentError) { @rtr.color = -1.1 }
    assert_nothing_raised { @rtr.color = "FF00FF00" }
    assert_equal(@rtr.color.rgb, "FF00FF00")
  end

  def test_scheme
    assert_raise(ArgumentError) { @rtr.scheme = -1.1 }
    assert_nothing_raised { @rtr.scheme = :major }
    assert_equal(@rtr.scheme, :major)
  end

  def test_vertAlign
    assert_raise(ArgumentError) { @rtr.vertAlign = -1.1 }
    assert_nothing_raised { @rtr.vertAlign = :baseline }
    assert_equal(@rtr.vertAlign, :baseline)
  end

  def test_sz
    assert_raise(ArgumentError) { @rtr.sz = -1.1 }
    assert_nothing_raised { @rtr.sz = 12 }
    assert_equal(@rtr.sz, 12)
  end

  def test_extend
    assert_raise(ArgumentError) { @rtr.extend = -1.1 }
    assert_nothing_raised { @rtr.extend = false }
    assert_equal(@rtr.extend, false)
  end

  def test_condense
    assert_raise(ArgumentError) { @rtr.condense = -1.1 }
    assert_nothing_raised { @rtr.condense = false }
    assert_equal(@rtr.condense, false)
  end

  def test_shadow
    assert_raise(ArgumentError) { @rtr.shadow = -1.1 }
    assert_nothing_raised { @rtr.shadow = false }
    assert_equal(@rtr.shadow, false)
  end

  def test_outline
    assert_raise(ArgumentError) { @rtr.outline = -1.1 }
    assert_nothing_raised { @rtr.outline = false }
    assert_equal(@rtr.outline, false)
  end

  def test_strike
    assert_raise(ArgumentError) { @rtr.strike = -1.1 }
    assert_nothing_raised { @rtr.strike = false }
    assert_equal(@rtr.strike, false)
  end

  def test_u
    @c.type = :string
    assert_raise(ArgumentError) { @c.u = -1.1 }
    assert_nothing_raised { @c.u = :single }
    assert_equal(@c.u, :single)
    doc = Nokogiri::XML(@c.to_xml_string(1,1))
    assert(doc.xpath('//u[@val="single"]'))
  end

  def test_i
    assert_raise(ArgumentError) { @c.i = -1.1 }
    assert_nothing_raised { @c.i = false }
    assert_equal(@c.i, false)
  end

  def test_rFont
    assert_raise(ArgumentError) { @c.font_name = -1.1 }
    assert_nothing_raised { @c.font_name = "Arial" }
    assert_equal(@c.font_name, "Arial")
  end

  def test_charset
    assert_raise(ArgumentError) { @c.charset = -1.1 }
    assert_nothing_raised { @c.charset = 1 }
    assert_equal(@c.charset, 1)
  end

  def test_family
    assert_raise(ArgumentError) { @c.family = -1.1 }
    assert_nothing_raised { @c.family = 5 }
    assert_equal(@c.family, 5)
  end

  def test_b
    assert_raise(ArgumentError) { @c.b = -1.1 }
    assert_nothing_raised { @c.b = false }
    assert_equal(@c.b, false)
  end
  
  def test_multiline_autowidth
    wrap = @p.workbook.styles.add_style({:alignment => {:wrap_text => true}})
    awtr = Axlsx::RichTextRun.new('I\'m bold' + "\n", :b => true)
    rt = Axlsx::RichText.new 
    rt.runs << awtr
    @ws.add_row [rt], :style => wrap
    ar = [0]
    awtr.autowidth(ar)
    assert_equal(ar.length, 2)
    assert_equal(ar.last, 0)
  end
  
  def test_to_xml
    schema = Nokogiri::XML::Schema(File.open(Axlsx::SML_XSD))
    doc = Nokogiri::XML(@ws.to_xml_string)
    errors = []
    schema.validate(doc).each do |error|
      puts error.message
      errors.push error
    end
    assert(errors.empty?, "error free validation")
    
    assert(doc.xpath('//rPr/b[@val=1]'))
    assert(doc.xpath('//rPr/i[@val=0]'))
    assert(doc.xpath('//rPr/b[@val=0]'))
    assert(doc.xpath('//rPr/i[@val=1]'))
    assert(doc.xpath('//is//t[contains(text(), "hihihi")]'))
    assert(doc.xpath('//is//t[contains(text(), "hihi2hi2")]'))
  end
end
