require 'tc_helper.rb'

class TestComments < Test::Unit::TestCase
  def setup
    p = Axlsx::Package.new
    wb = p.workbook
    @ws = wb.add_worksheet
    @c1 = @ws.add_comment :ref => 'A1', :text => 'penut machine', :author => 'crank'
    @c2 = @ws.add_comment :ref => 'C3', :text => 'rust bucket', :author => 'PO'
  end

  def test_initialize
    assert_raise(ArgumentError) { Axlsx::Comments.new }
    assert(@ws.comments.vml_drawing.is_a?(Axlsx::VmlDrawing))
  end

  def test_add_comment
    assert_equal(@ws.comments.size, 2)
    assert_raise(ArgumentError) { @ws.comments.add_comment() }
    assert_raise(ArgumentError) { @ws.comments.add_comment(:text => 'Yes We Can', :ref => 'A1') }
    assert_raise(ArgumentError) { @ws.comments.add_comment(:author => 'bob', :ref => 'A1') }
    assert_raise(ArgumentError) { @ws.comments.add_comment(:author => 'bob', :text => 'Yes We Can')}
    assert_nothing_raised { @ws.comments.add_comment(:author => 'bob', :text => 'Yes We Can', :ref => 'A1') }
    assert_equal(@ws.comments.size, 3)
  end
  def test_authors
    assert_equal(@ws.comments.authors.size, @ws.comments.size)
    @ws.add_comment(:text => 'Yes We Can!', :author => 'bob', :ref => 'F1')
    assert_equal(@ws.comments.authors.size, 3)
    @ws.add_comment(:text => 'Yes We Can!', :author => 'bob', :ref => 'F1')
    assert_equal(@ws.comments.authors.size, 3, 'only unique authors are returned') 
  end
  def test_pn
    assert_equal(@ws.comments.pn, Axlsx::COMMENT_PN % (@ws.index+1).to_s)
  end

  def test_index
    assert_equal(@ws.index, @ws.comments.index)
  end

  def test_to_xml_string
    doc = Nokogiri::XML(@ws.comments.to_xml_string)
    schema = Nokogiri::XML::Schema(File.open(Axlsx::SML_XSD))
    errors = []
    schema.validate(doc).each do |error|
      errors << error
    end
    assert_equal(0, errors.length)

    # TODO figure out why these xpath expressions dont work!
    # assert(doc.xpath("//comments"))
    # assert_equal(doc.xpath("//xmlns:author").size, @ws.comments.authors.size)
    # assert_equal(doc.xpath("//comment").size, @ws.comments.size)
  end
end


