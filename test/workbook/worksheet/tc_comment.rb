require 'tc_helper.rb'

class TestComment < Test::Unit::TestCase
  def setup
    p = Axlsx::Package.new
    wb = p.workbook
    @ws = wb.add_worksheet
    @c1 = @ws.add_comment :ref => 'A1', :text => 'text with special char <', :author => 'author with special char <', :visible => false
    @c2 = @ws.add_comment :ref => 'C3', :text => 'rust bucket', :author => 'PO'
  end

  def test_initailize
    assert_raise(ArgumentError) { Axlsx::Comment.new }
  end

  def test_author
    assert(@c1.author == 'author with special char <')
    assert(@c2.author == 'PO')
  end

  def test_text
    assert(@c1.text == 'text with special char <')
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

  def test_to_xml_string
    doc = Nokogiri::XML(@c1.to_xml_string)
    assert_equal(doc.xpath("//comment[@ref='#{@c1.ref}']").size, 1)
    assert_equal(doc.xpath("//comment[@authorId='#{@c1.author_index.to_s}']").size, 1)
    assert_equal(doc.xpath("//t[text()='#{@c1.author}:\n']").size, 1)
    assert_equal(doc.xpath("//t[text()='#{@c1.text}']").size, 1)
  end

  def test_comment_text_contain_author_and_text
    comment = @ws.add_comment :ref => 'C4', :text => 'some text', :author => 'Bob'
    doc = Nokogiri::XML(comment.to_xml_string)
    assert_equal("Bob:\nsome text", doc.xpath("//comment/text").text)
  end

  def test_comment_text_does_not_contain_stray_colon_if_author_blank
    comment = @ws.add_comment :ref => 'C5', :text => 'some text', :author => ''
    doc = Nokogiri::XML(comment.to_xml_string)
    assert_equal("some text", doc.xpath("//comment/text").text)
  end
end

