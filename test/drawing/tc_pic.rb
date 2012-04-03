require 'tc_helper.rb'

class TestPic < Test::Unit::TestCase

  def setup
    @p = Axlsx::Package.new
    ws = @p.workbook.add_worksheet
    @test_img =  File.dirname(__FILE__) + "/../../examples/image1.jpeg"
    @image = ws.add_image :image_src => @test_img
  end

  def teardown
  end

  def test_initialization
    assert_equal(@p.workbook.images.first, @image)
    assert_equal(@image.image_src, @test_img)
  end

  def test_hyperlink
    assert_equal(@image.hyperlink, nil)
    @image.hyperlink = "http://axlsx.blogspot.com"
    assert_equal(@image.hyperlink.href, "http://axlsx.blogspot.com")
  end

  def test_name
    assert_raise(ArgumentError) { @image.name = 49 }
    assert_nothing_raised { @image.name = "unknown" }
    assert_equal(@image.name, "unknown")
  end

  def test_start_at
    assert_raise(ArgumentError) { @image.start_at "a", 1 }
    assert_nothing_raised { @image.start_at 6, 7 }
    assert_equal(@image.anchor.from.col, 6)
    assert_equal(@image.anchor.from.row, 7)
  end

  def test_width
    assert_raise(ArgumentError) { @image.width = "a" }
    assert_nothing_raised { @image.width = 600 }
    assert_equal(@image.width, 600)
  end

  def test_height
    assert_raise(ArgumentError) { @image.height = "a" }
    assert_nothing_raised { @image.height = 600 }
    assert_equal(600, @image.height)
  end

  def test_image_src
    assert_raise(ArgumentError) { @image.image_src = 49 }
    assert_raise(ArgumentError) { @image.image_src = 'Unknown' }
    assert_raise(ArgumentError) { @image.image_src = __FILE__ }
    assert_nothing_raised { @image.image_src = @test_img }
    assert_equal(@image.image_src, @test_img)
  end

  def test_descr
    assert_raise(ArgumentError) { @image.descr = 49 }
    assert_nothing_raised { @image.descr = "test" }
    assert_equal(@image.descr, "test")
  end

  def test_to_xml
    schema = Nokogiri::XML::Schema(File.open(Axlsx::DRAWING_XSD))
    doc = Nokogiri::XML(@image.anchor.drawing.to_xml_string)
    errors = []
    schema.validate(doc).each do |error|
      errors.push error
      puts error.message
    end
    assert(errors.empty?, "error free validation")
  end

end
