require 'tc_helper.rb'

class TestComment < Test::Unit::TestCase
  def setup
    p = Axlsx::Package.new
    wb = p.workbook
    @ws = wb.add_worksheet
    @c1 = @ws.add_comment :ref => 'A1', :text => 'penut machine', :author => 'crank', :visible => false
    @c2 = @ws.add_comment :ref => 'C3', :text => 'rust bucket', :author => 'PO'
  end

  def test_initailize
    assert_raise(ArgumentError) { Axlsx::Comment.new }
  end

  def test_author
    assert(@c1.author == 'crank')
    assert(@c2.author == 'PO')
  end

  def test_text
    assert(@c1.text == 'penut machine')
    assert(@c2.text == 'rust bucket')
  end

  def test_author_index
    assert_equal(@c1.author_index, 1)
    assert_equal(@c2.author_index, 0)
  end

  def test_visible
    assert_equal(false, @c1.visible)
    assert_equal(true, @c2.visible)
  end
  def test_ref
    assert(@c1.ref == 'A1')
    assert(@c2.ref == 'C3')
  end

  def test_vml_shape
    pos = Axlsx::name_to_indices(@c1.ref)
    assert(@c1.vml_shape.is_a?(Axlsx::VmlShape))
    assert(@c1.vml_shape.column == pos[0])
    assert(@c1.vml_shape.row == pos[1])
    assert(@c1.vml_shape.row == pos[1])
    assert_equal(pos[0], @c1.vml_shape.left_column)
    assert(@c1.vml_shape.top_row == pos[1])
    assert_equal(pos[0] + 2 , @c1.vml_shape.right_column)
    assert(@c1.vml_shape.bottom_row == pos[1]+4)
  end

  def to_xml_string
    doc = Nokogiri::XML(@c1.to_xml_string)
    assert_equal(doc.xpath("//comment[@ref='#{@c1.ref}']").size, 1)
    assert_equal(doc.xpath("//comment[@authorId='#{@c1.author_index.to}']").size, 1)
    assert_equal(doc.xpath("//t[text()='#{@c1.author}']").size, 1)
    assert_equal(doc.xpath("//t[text()='#{@c1.text}']").size, 1)
  end

end

