require 'tc_helper.rb'

class TestVmlDrawing < Test::Unit::TestCase

  def setup
    p = Axlsx::Package.new
    wb = p.workbook
    @ws = wb.add_worksheet
    @ws.add_comment :ref => 'C3', :text => 'rust bucket', :author => 'PO'
    @vml_drawing = @ws.comments.vml_drawing
  end

  def test_initialize
    assert_raise(ArgumentError) { Axlsx::VmlDrawing.new }
  end

  def test_to_xml_string
    str = '<?xml version="1.0" encoding="UTF-8"?>'
    str << '<c:chartSpace xmlns:c="' << Axlsx::XML_NS_C << '">'
    str << @vml_drawing.to_xml_string(0)
    doc = Nokogiri::XML(str)
  end

end
