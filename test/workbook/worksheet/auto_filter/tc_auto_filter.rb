require 'tc_helper.rb'

class TestAutoFilter < Test::Unit::TestCase

  def setup
    ws = Axlsx::Package.new.workbook.add_worksheet
    3.times { |index| ws.add_row [1*index,2*index,3*index] }
    @auto_filter = ws.auto_filter
    @auto_filter.range = 'A1:C3'
    @auto_filter.add_column 0, :filters, :filter_items => [1]
  end

  def test_defined_name
    assert_equal("'Sheet1'!$A$1:$C$3", @auto_filter.defined_name)
  end

  def test_to_xml_string
    doc = Nokogiri::XML(@auto_filter.to_xml_string)
    assert(doc.xpath("autoFilter[@ref='#{@auto_filter.range}']"))
  end

  def test_columns
    assert @auto_filter.columns.is_a?(Axlsx::SimpleTypedList)
    assert_equal @auto_filter.columns.allowed_types, [Axlsx::FilterColumn]
  end

  def test_add_column
    @auto_filter.add_column(0, :filters) do |column|
      assert column.is_a? FilterColumn
    end
  end

  def test_applya
    assert_equal nil, @auto_filter.worksheet.rows.last.hidden
    @auto_filter.apply
    assert_equal true, @auto_filter.worksheet.rows.last.hidden
  end
end
