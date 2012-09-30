require 'tc_helper'

class TestSheetCalcPr < Test::Unit::TestCase

  def setup
    @sheet_calc_pr = Axlsx::SheetCalcPr.new(:full_calc_on_load => false)
  end

  def test_full_calc_on_load
    assert_equal false, @sheet_calc_pr.full_calc_on_load
    assert Axlsx::SheetCalcPr.new.full_calc_on_load
  end

  def test_to_xml_string
    doc = Nokogiri::XML(@sheet_calc_pr.to_xml_string)
    assert_equal 1, doc.xpath('//sheetCalcPr[@fullCalcOnLoad="false"]').size
  end
end
