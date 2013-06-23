require 'tc_helper.rb'

class TestAxes < Test::Unit::TestCase
  def setup
    @axes = Axlsx::Axes.new(:val_axis => Axlsx::ValAxis, :cat_axis => Axlsx::CatAxis)
  end

  def test_to_xml_string_just_ids
    str  = '<?xml version="1.0" encoding="UTF-8"?>'
    str << '<c:chartSpace xmlns:c="' << Axlsx::XML_NS_C << '" xmlns:a="' << Axlsx::XML_NS_A << '">'
    @axes.to_xml_string(str, :ids => true)
    cat_axis_position = str.index(@axes[:cat_axis].id.to_s)
    val_axis_position = str.index(@axes[:val_axis].id.to_s)
    assert(cat_axis_position < val_axis_position, "cat_axis must occur earlier than val_axis in the XML")
  end
end