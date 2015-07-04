require 'tc_helper.rb'


class TestSheetPr < Test::Unit::TestCase

  def setup
    worksheet = Axlsx::Package.new.workbook.add_worksheet
    @options = {
      :sync_horizontal => false,
      :sync_vertical => false,
      :transition_evaluation => true,
      :transition_entry => true,
      :published => false,
      :filter_mode => true,
      :enable_format_conditions_calculation => false,
      :code_name => '007',
      :sync_ref => 'foo',
      :tab_color => 'FFFF6666'
    }
    @sheet_pr = Axlsx::SheetPr.new(worksheet, @options)
  end

  def test_initialization
    @options.each do |key, value|
      if key==:tab_color
        stored_value = @sheet_pr.send(key)
        assert_equal Axlsx::Color, stored_value.class
        assert_equal value, stored_value.rgb
      else
        assert_equal value, @sheet_pr.send(key)
      end
    end
  end

  def test_to_xml_string
    doc = Nokogiri::XML(@sheet_pr.to_xml_string)
    assert_equal(doc.xpath("//sheetPr[@syncHorizontal='0']").size, 1)
    assert_equal(doc.xpath("//sheetPr[@syncVertical='0']").size, 1)
    assert_equal(doc.xpath("//sheetPr[@transitionEvaluation='1']").size, 1)
    assert_equal(doc.xpath("//sheetPr[@transitionEntry='1']").size, 1)
    assert_equal(doc.xpath("//sheetPr[@published='0']").size, 1)
    assert_equal(doc.xpath("//sheetPr[@filterMode='1']").size, 1)
    assert_equal(doc.xpath("//sheetPr[@enableFormatConditionsCalculation='0']").size, 1)
    assert_equal(doc.xpath("//sheetPr[@codeName='007']").size, 1)
    assert_equal(doc.xpath("//sheetPr[@syncRef='foo']").size, 1)
    assert_equal(doc.xpath("//sheetPr/tabColor[@rgb='FFFF6666']").size, 1)
    assert_equal(doc.xpath("//sheetPr/pageSetUpPr[@fitToPage='0']").size, 1)
  end
end
