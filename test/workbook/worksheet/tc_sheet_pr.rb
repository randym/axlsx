require 'tc_helper.rb'


class TestSheetPr < Test::Unit::TestCase

  def setup
    worksheet = Axlsx::Package.new.workbook.add_worksheet
    @options = { 
      :sync_horizontal => false,
      :sync_vertical => false,
      :transtion_evaluation => true,
      :transition_entry => true,
      :published => false,
      :filter_mode => true,
      :enable_format_conditions_calculation => false,
      :code_name => '007',
      :sync_ref => 'foo'
    }
    @sheet_pr = Axlsx::SheetPr.new(worksheet, @options)
    @sheet_pr.tab_color = "FFFF6666"
  end

  def test_initialization
    @options.each do |key, value|
      assert_equal value, @sheet_pr.send(key)
    end
  end

  def test_tab_color
    assert_equal "FFFF6666", @sheet_pr.tab_color
  end
end
