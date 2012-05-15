require 'tc_helper.rb'

class TestTableStyle < Test::Unit::TestCase

  def setup
    @item = Axlsx::TableStyle.new "fisher"
  end

  def teardown
  end

  def test_initialiation
    assert_equal(@item.name, "fisher")
    assert_equal(@item.pivot, nil)
    assert_equal(@item.table, nil)
    ts = Axlsx::TableStyle.new 'price', :pivot => true, :table => true
    assert_equal(ts.name, 'price')
    assert_equal(ts.pivot, true)
    assert_equal(ts.table, true)
  end

  def test_name
    assert_raise(ArgumentError) { @item.name = -1.1 }
    assert_nothing_raised { @item.name = "lovely table style" }
    assert_equal(@item.name, "lovely table style")
  end

  def test_pivot
    assert_raise(ArgumentError) { @item.pivot = -1.1 }
    assert_nothing_raised { @item.pivot = true }
    assert_equal(@item.pivot, true)
  end

  def test_table
    assert_raise(ArgumentError) { @item.table = -1.1 }
    assert_nothing_raised { @item.table = true }
    assert_equal(@item.table, true)
  end

  def test_to_xml_string
    doc = Nokogiri::XML(@item.to_xml_string)
    assert(doc.xpath("//tableStyle[@name='#{@item.name}']"))
  end
end
