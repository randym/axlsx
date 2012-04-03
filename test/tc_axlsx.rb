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

  def test_cell_range
    #To do
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

end
