require 'tc_helper.rb'

class TestOneCellAnchor < Test::Unit::TestCase

  def setup
    @p = Axlsx::Package.new
    @ws = @p.workbook.add_worksheet
    @test_img =  File.dirname(__FILE__) + "/../../examples/image1.jpeg"
    @image = @ws.add_image :image_src => @test_img
    @anchor = @image.anchor
  end

  def teardown
  end

  def test_initialization
    assert(@anchor.from.col == 0)
    assert(@anchor.from.row == 0)
    assert(@anchor.width == 0)
    assert(@anchor.height == 0)
  end

  def test_from
    assert(@anchor.from.is_a?(Axlsx::Marker))
  end

  def test_object
    assert(@anchor.object.is_a?(Axlsx::Pic))
  end

  def test_index
    assert_equal(@anchor.index, @anchor.drawing.anchors.index(@anchor))
  end

  def test_width
    assert_raise(ArgumentError) { @anchor.width = "a" }
    assert_nothing_raised { @anchor.width = 600 }
    assert_equal(@anchor.width, 600)
  end

  def test_height
    assert_raise(ArgumentError) { @anchor.height = "a" }
    assert_nothing_raised { @anchor.height = 400 }
    assert_equal(400, @anchor.height)
  end

  def test_ext
    ext = @anchor.send(:ext)
    assert_equal(ext[:cx], (@anchor.width * 914400 / 96))
    assert_equal(ext[:cy], (@anchor.height * 914400 / 96))
  end

  def test_options
    assert_raise(ArgumentError, 'invalid start_at') { @ws.add_image :image_src=>@test_img, :start_at=>[1] }
    i = @ws.add_image :image_src=>@test_img, :start_at => [1,2], :width=>100, :height=>200, :name=>"someimage", :descr=>"a neat image"

    assert_equal("a neat image", i.descr)
    assert_equal("someimage", i.name)
    assert_equal(200, i.height)
    assert_equal(100, i.width)
    assert_equal(1, i.anchor.from.col)
    assert_equal(2, i.anchor.from.row)
    assert_equal(@test_img, i.image_src)
  end

end
