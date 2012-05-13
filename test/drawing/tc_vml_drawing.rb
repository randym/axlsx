require 'tc_helper.rb'

class TestVmlDrawing < Test::Unit::TestCase

  def setup
    p = Axlsx::Package.new
    wb = p.workbook
    @ws = wb.add_worksheet
    @ws.add_comment :ref => 'A1', :text => 'penut machine', :author => 'crank'
    @ws.add_comment :ref => 'C3', :text => 'rust bucket', :author => 'PO'
    @vml_drawing = @ws.comments.vml_drawing
  end

  def test_initialize
    assert_raise(ArgumentError) { Axlsx::VmlDrawing.new }
  end

  def test_to_xml_string
    str = @vml_drawing.to_xml_string()
    doc = Nokogiri::XML(str)
    assert_equal(doc.xpath("//v:shape").size, 2)
    assert(doc.xpath("//o:idmap[@o:data='#{@ws.index+1}']"))
  end

end
