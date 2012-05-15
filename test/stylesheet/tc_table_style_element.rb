require 'tc_helper.rb'

class TestTableStyleElement < Test::Unit::TestCase

  def setup
    @item = Axlsx::TableStyleElement.new
  end

  def teardown
  end

  def test_initialiation
    assert_equal(@item.type, nil)
    assert_equal(@item.size, nil)
    assert_equal(@item.dxfId, nil)
    options = { :type => :headerRow, :size => 10, :dxfId => 1 }

    tse = Axlsx::TableStyleElement.new options
    options.each { |key, value| assert_equal(tse.send(key.to_sym), value) }
  end

  def test_type
    assert_raise(ArgumentError) { @item.type = -1.1 }
    assert_nothing_raised { @item.type = :blankRow }
    assert_equal(@item.type, :blankRow)
  end

  def test_size
    assert_raise(ArgumentError) { @item.size = -1.1 }
    assert_nothing_raised { @item.size = 2 }
    assert_equal(@item.size, 2)
  end

  def test_dxfId
    assert_raise(ArgumentError) { @item.dxfId = -1.1 }
    assert_nothing_raised { @item.dxfId = 7 }
    assert_equal(@item.dxfId, 7)
  end
  
  def test_to_xml_string
    doc = Nokogiri::XML(@item.to_xml_string)
   @item.type = :headerRow
    assert(doc.xpath("//tableStyleElement[@type='#{@item.type.to_s}']"))
  end
end
