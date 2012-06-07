require 'tc_helper.rb'

class TestPic < Test::Unit::TestCase

  def setup
    @p = Axlsx::Package.new
    ws = @p.workbook.add_worksheet
    @test_img =  File.dirname(__FILE__) + "/../../examples/image1.jpeg"
    @image = ws.add_image :image_src => @test_img, :hyperlink => 'https://github.com/randym', :tooltip => "What's up doc?"
  end

  def teardown
  end

  def test_initialization
    assert_equal(@p.workbook.images.first, @image)
    assert_equal(@image.file_name, 'image1.jpeg')
    assert_equal(@image.image_src, @test_img)
  end

  def test_anchor_swapping
    #swap from one cell to two cell when end_at is specified
    assert(@image.anchor.is_a?(Axlsx::OneCellAnchor))
    start_at = @image.anchor.from 
    @image.end_at 10,5
    assert(@image.anchor.is_a?(Axlsx::TwoCellAnchor))
    assert_equal(start_at.col, @image.anchor.from.col)
    assert_equal(start_at.row, @image.anchor.from.row)
    assert_equal(10,@image.anchor.to.col)
    assert_equal(5, @image.anchor.to.row)
  
    #swap from two cell to one cell when width or height are specified
    @image.width = 200
    assert(@image.anchor.is_a?(Axlsx::OneCellAnchor))
    assert_equal(start_at.col, @image.anchor.from.col)
    assert_equal(start_at.row, @image.anchor.from.row)
    assert_equal(200, @image.width)
  
  end
  def test_hyperlink
    assert_equal(@image.hyperlink.href, "https://github.com/randym")
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
