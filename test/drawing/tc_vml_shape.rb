require 'tc_helper.rb'

class TestVmlShape < Test::Unit::TestCase

  def setup
    p = Axlsx::Package.new
    wb = p.workbook
    @ws = wb.add_worksheet
    @ws.add_comment :ref => 'A1', :text => 'penut machine', :author => 'crank', :visible => true
    @ws.add_comment :ref => 'C3', :text => 'rust bucket', :author => 'PO', :visible => false
    @comments = @ws.comments
  end

  def test_initialize
    assert_raise(ArgumentError) { Axlsx::VmlDrawing.new }
  end

  def test_row
    shape = @comments.first.vml_shape
    assert_equal(shape.row, 0)
    shape = @comments.last.vml_shape
    assert_equal(shape.row, 2)
  end

  def test_column
    shape = @comments.first.vml_shape
    assert_equal(shape.column, 0)
    shape = @comments.last.vml_shape
    assert_equal(shape.column, 2)
  end

  def test_left_column
    shape = @comments.first.vml_shape
    shape.left_column = 3
    assert(shape.left_column == 3)
    assert_raise(ArgumentError) { shape.left_column = [] }
  end

  def test_left_offset
    shape = @comments.first.vml_shape
    shape.left_offset = 3
    assert(shape.left_offset == 3)
    assert_raise(ArgumentError) { shape.left_offset = [] }
  end

  def test_right_column
    shape = @comments.first.vml_shape
    shape.right_column = 3
    assert(shape.right_column == 3)
    assert_raise(ArgumentError) { shape.right_column = [] }
  end

  def test_right_offset
    shape = @comments.first.vml_shape
    shape.right_offset = 3
    assert(shape.right_offset == 3)
    assert_raise(ArgumentError) { shape.right_offset = [] }
  end

  def test_top_offset
    shape = @comments.first.vml_shape
    shape.top_offset = 3
    assert(shape.top_offset == 3)
    assert_raise(ArgumentError) { shape.top_offset = [] }
  end

  def test_bottom_offset
    shape = @comments.first.vml_shape
    shape.bottom_offset = 3
    assert(shape.bottom_offset == 3)
    assert_raise(ArgumentError) { shape.bottom_offset = [] }
  end

  def test_bottom_row
    shape = @comments.first.vml_shape
    shape.bottom_row = 3
    assert(shape.bottom_row == 3)
    assert_raise(ArgumentError) { shape.bottom_row = [] }
  end

  def test_top_row
    shape = @comments.first.vml_shape
    shape.top_row = 3
    assert(shape.top_row == 3)
    assert_raise(ArgumentError) { shape.top_row = [] }
  end
  def test_visible
    shape = @comments.first.vml_shape
    shape.visible = false
    assert(shape.visible == false)
    assert_raise(ArgumentError) { shape.visible = 'foo' }
  end
  def test_to_xml_string
    str = @comments.vml_drawing.to_xml_string()
    doc = Nokogiri::XML(str)
    assert_equal(doc.xpath("//v:shape").size, 2)
    assert_equal(1, doc.xpath("//x:Visible").size, 'ClientData/x:Visible element rendering')
    @comments.each do |comment|
      shape = comment.vml_shape
      assert(doc.xpath("//v:shape/x:ClientData/x:Row[text()='#{shape.row}']").size == 1)
      assert(doc.xpath("//v:shape/x:ClientData/x:Column[text()='#{shape.column}']").size == 1)
      assert(doc.xpath("//v:shape/x:ClientData/x:Anchor[text()='#{shape.left_column}, #{shape.left_offset}, #{shape.top_row}, #{shape.top_offset}, #{shape.right_column}, #{shape.right_offset}, #{shape.bottom_row}, #{shape.bottom_offset}']").size == 1)
    end
  end

end
