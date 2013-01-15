require 'tc_helper.rb'

class TestAxlsx < Test::Unit::TestCase

  def setup_wide
    @wide_test_points = { "A3" =>      0,
      "Z3"    =>                      25,
      "B3"    =>                       1,
      "AA3"   =>             1 * 26 +  0,
      "AAA3"  => 1 * 26**2 + 1 * 26 +  0,
      "AAZ3"  => 1 * 26**2 + 1 * 26 + 25,
      "ABA3"  => 1 * 26**2 + 2 * 26 +  0,
      "BZU3"  => 2 * 26**2 + 26 * 26 + 20
    }
  end

  def test_cell_range_empty_if_no_cell
    assert_equal(Axlsx.cell_range([]), "")
  end

  def test_do_not_trust_input_by_default
    assert_equal false, Axlsx.trust_input
  end


  def test_trust_input_can_be_set_to_true
    Axlsx.trust_input = true
    assert_equal true, Axlsx.trust_input
  end
  def test_cell_range_relative
    p = Axlsx::Package.new
    ws = p.workbook.add_worksheet
    row = ws.add_row
    c1 = row.add_cell
    c2 = row.add_cell
    assert_equal(Axlsx.cell_range([c2, c1], false), "A1:B1")
  end

  def test_cell_range_absolute
    p = Axlsx::Package.new
    ws = p.workbook.add_worksheet :name => "Sheet <'>\" 1"
    row = ws.add_row
    c1 = row.add_cell
    c2 = row.add_cell
    assert_equal(Axlsx.cell_range([c2, c1], true), "'Sheet &lt;''&gt;&quot; 1'!$A$1:$B$1")
  end

  def test_name_to_indices
    setup_wide
    @wide_test_points.each do |key, value|
      assert_equal(Axlsx.name_to_indices(key), [value,2])
    end
  end

  def test_col_ref
    setup_wide
    @wide_test_points.each do |key, value|
      assert_equal(Axlsx.col_ref(value), key.gsub(/\d+/, ''))
    end
  end

  def test_cell_r
    # todo
  end

  def test_range_to_a
    assert_equal([['A1', 'B1', 'C1']],                         Axlsx::range_to_a('A1:C1'))
    assert_equal([['A1', 'B1', 'C1'], ['A2', 'B2', 'C2']],     Axlsx::range_to_a('A1:C2'))
    assert_equal([['Z5', 'AA5', 'AB5'], ['Z6', 'AA6', 'AB6']], Axlsx::range_to_a('Z5:AB6'))
  end

end
