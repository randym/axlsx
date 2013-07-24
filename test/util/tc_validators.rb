require 'tc_helper.rb'
class TestValidators < Test::Unit::TestCase
  def setup
  end
  def teardown
  end

  def test_validators
    #unsigned_int
    assert_nothing_raised { Axlsx.validate_unsigned_int 1 }
    assert_nothing_raised { Axlsx.validate_unsigned_int(+1) }
    assert_raise(ArgumentError) { Axlsx.validate_unsigned_int(-1)}
    assert_raise(ArgumentError) { Axlsx.validate_unsigned_int('1') }

    #int
    assert_nothing_raised { Axlsx.validate_int(1) }
    assert_nothing_raised { Axlsx.validate_int(-1) }
    assert_raise(ArgumentError) { Axlsx.validate_int('a')}
    assert_raise(ArgumentError) { Axlsx.validate_int(Array) }

    #boolean (as 0 or 1, :true, :false, true, false, or "true," "false")
    [0,1,:true, :false, true, false, "true", "false"].each do |v|
      assert_nothing_raised { Axlsx.validate_boolean 0 }
    end
    assert_raise(ArgumentError) { Axlsx.validate_boolean 2 }

    #string
    assert_nothing_raised { Axlsx.validate_string "1" }
    assert_raise(ArgumentError) { Axlsx.validate_string 2 }
    assert_raise(ArgumentError) { Axlsx.validate_string false }

    #float
    assert_nothing_raised { Axlsx.validate_float 1.0 }
    assert_raise(ArgumentError) { Axlsx.validate_float 2 }
    assert_raise(ArgumentError) { Axlsx.validate_float false }

    #pattern_type
    assert_nothing_raised { Axlsx.validate_pattern_type :none }
    assert_raise(ArgumentError) { Axlsx.validate_pattern_type "none" }
    assert_raise(ArgumentError) { Axlsx.validate_pattern_type "crazy_pattern" }
    assert_raise(ArgumentError) { Axlsx.validate_pattern_type false }

    #gradient_type
    assert_nothing_raised { Axlsx.validate_gradient_type :path }
    assert_raise(ArgumentError) { Axlsx.validate_gradient_type nil }
    assert_raise(ArgumentError) { Axlsx.validate_gradient_type "fractal" }
    assert_raise(ArgumentError) { Axlsx.validate_gradient_type false }

    #horizontal alignment
    assert_nothing_raised { Axlsx.validate_horizontal_alignment :general }
    assert_raise(ArgumentError) { Axlsx.validate_horizontal_alignment nil }
    assert_raise(ArgumentError) { Axlsx.validate_horizontal_alignment "wavy" }
    assert_raise(ArgumentError) { Axlsx.validate_horizontal_alignment false }

    #vertical alignment
    assert_nothing_raised { Axlsx.validate_vertical_alignment :top }
    assert_raise(ArgumentError) { Axlsx.validate_vertical_alignment nil }
    assert_raise(ArgumentError) { Axlsx.validate_vertical_alignment "dynamic" }
    assert_raise(ArgumentError) { Axlsx.validate_vertical_alignment false }

    #contentType
    assert_nothing_raised { Axlsx.validate_content_type Axlsx::WORKBOOK_CT }
    assert_raise(ArgumentError) { Axlsx.validate_content_type nil }
    assert_raise(ArgumentError) { Axlsx.validate_content_type "http://some.url" }
    assert_raise(ArgumentError) { Axlsx.validate_content_type false }

    #relationshipType
    assert_nothing_raised { Axlsx.validate_relationship_type Axlsx::WORKBOOK_R }
    assert_raise(ArgumentError) { Axlsx.validate_relationship_type nil }
    assert_raise(ArgumentError) { Axlsx.validate_relationship_type "http://some.url" }
    assert_raise(ArgumentError) { Axlsx.validate_relationship_type false }

    #number_with_unit
    assert_nothing_raised { Axlsx.validate_number_with_unit "210mm" }
    assert_nothing_raised { Axlsx.validate_number_with_unit "8.5in" }
    assert_nothing_raised { Axlsx.validate_number_with_unit "29.7cm" }
    assert_nothing_raised { Axlsx.validate_number_with_unit "120pt" }
    assert_nothing_raised { Axlsx.validate_number_with_unit "0pc" }
    assert_nothing_raised { Axlsx.validate_number_with_unit "12.34pi" }
    assert_raise(ArgumentError) { Axlsx.validate_number_with_unit nil }
    assert_raise(ArgumentError) { Axlsx.validate_number_with_unit "210" }
    assert_raise(ArgumentError) { Axlsx.validate_number_with_unit 210 }
    assert_raise(ArgumentError) { Axlsx.validate_number_with_unit "mm" }
    assert_raise(ArgumentError) { Axlsx.validate_number_with_unit "-29cm" }

    #scale_10_400
    assert_nothing_raised { Axlsx.validate_scale_10_400 10 }
    assert_nothing_raised { Axlsx.validate_scale_10_400 100 }
    assert_nothing_raised { Axlsx.validate_scale_10_400 400 }
    assert_raise(ArgumentError) { Axlsx.validate_scale_10_400 9 }
    assert_raise(ArgumentError) { Axlsx.validate_scale_10_400 10.0 }
    assert_raise(ArgumentError) { Axlsx.validate_scale_10_400 400.1 }
    assert_raise(ArgumentError) { Axlsx.validate_scale_10_400 "99" }
    
    #scale_0_10_400
    assert_nothing_raised { Axlsx.validate_scale_0_10_400 0 }
    assert_nothing_raised { Axlsx.validate_scale_0_10_400 10 }
    assert_nothing_raised { Axlsx.validate_scale_0_10_400 100 }
    assert_nothing_raised { Axlsx.validate_scale_0_10_400 400 }
    assert_raise(ArgumentError) { Axlsx.validate_scale_0_10_400 9 }
    assert_raise(ArgumentError) { Axlsx.validate_scale_0_10_400 10.0 }
    assert_raise(ArgumentError) { Axlsx.validate_scale_0_10_400 400.1 }
    assert_raise(ArgumentError) { Axlsx.validate_scale_0_10_400 "99" }
    
    #page_orientation
    assert_nothing_raised { Axlsx.validate_page_orientation :default }
    assert_nothing_raised { Axlsx.validate_page_orientation :landscape }
    assert_nothing_raised { Axlsx.validate_page_orientation :portrait }
    assert_raise(ArgumentError) { Axlsx.validate_page_orientation nil }
    assert_raise(ArgumentError) { Axlsx.validate_page_orientation 1 }
    assert_raise(ArgumentError) { Axlsx.validate_page_orientation "landscape" }
    
    #data_validation_error_style
    [:information, :stop, :warning].each do |sym|
      assert_nothing_raised { Axlsx.validate_data_validation_error_style sym }
    end
    assert_raise(ArgumentError) { Axlsx.validate_data_validation_error_style :other_symbol }
    assert_raise(ArgumentError) { Axlsx.validate_data_validation_error_style 'warning' }
    assert_raise(ArgumentError) { Axlsx.validate_data_validation_error_style 0 }
    
    #data_validation_operator
    [:lessThan, :lessThanOrEqual, :equal, :notEqual, :greaterThanOrEqual, :greaterThan, :between, :notBetween].each do |sym|
      assert_nothing_raised { Axlsx.validate_data_validation_operator sym }
    end
    assert_raise(ArgumentError) { Axlsx.validate_data_validation_error_style :other_symbol }
    assert_raise(ArgumentError) { Axlsx.validate_data_validation_error_style 'lessThan' }
    assert_raise(ArgumentError) { Axlsx.validate_data_validation_error_style 0 }
    
    #data_validation_type
    [:custom, :data, :decimal, :list, :none, :textLength, :time, :whole].each do |sym|
      assert_nothing_raised { Axlsx.validate_data_validation_type sym }
    end
    assert_raise(ArgumentError) { Axlsx.validate_data_validation_error_style :other_symbol }
    assert_raise(ArgumentError) { Axlsx.validate_data_validation_error_style 'decimal' }
    assert_raise(ArgumentError) { Axlsx.validate_data_validation_error_style 0 }
    
    #sheet_view_type
    [:normal, :page_break_preview, :page_layout].each do |sym|
      assert_nothing_raised { Axlsx.validate_sheet_view_type sym }
    end
    assert_raise(ArgumentError) { Axlsx.validate_data_validation_error_style :other_symbol }
    assert_raise(ArgumentError) { Axlsx.validate_data_validation_error_style 'page_layout' }
    assert_raise(ArgumentError) { Axlsx.validate_data_validation_error_style 0 }
    
    #active_pane_type
    [:bottom_left, :bottom_right, :top_left, :top_right].each do |sym|
      assert_nothing_raised { Axlsx.validate_pane_type sym }
    end
    assert_raise(ArgumentError) { Axlsx.validate_pane_type :other_symbol }
    assert_raise(ArgumentError) { Axlsx.validate_pane_type 'bottom_left' }
    assert_raise(ArgumentError) { Axlsx.validate_pane_type 0 }
    
    #split_state_type
    [:frozen, :frozen_split, :split].each do |sym|
      assert_nothing_raised { Axlsx.validate_split_state_type sym }
    end
    assert_raise(ArgumentError) { Axlsx.validate_split_state_type :other_symbol }
    assert_raise(ArgumentError) { Axlsx.validate_split_state_type 'frozen_split' }
    assert_raise(ArgumentError) { Axlsx.validate_split_state_type 0 }
  end
 
  def test_validate_integerish
    assert_raise(ArgumentError) { Axlsx.validate_integerish Axlsx }
    [1, 1.4, "a"].each { |test_value| assert_nothing_raised { Axlsx.validate_integerish test_value } }
  end
  def test_validate_family
    assert_raise(ArgumentError) { Axlsx.validate_family 0 }
    (1..5).each do |item|
      assert_nothing_raised { Axlsx.validate_family item }
    end
  end

  def test_validate_u
    assert_raise(ArgumentError) { Axlsx.validate_cell_u :hoge }
    [:none, :single, :double, :singleAccounting, :doubleAccounting].each do |sym|
      assert_nothing_raised { Axlsx.validate_cell_u sym }
    end
  end

  def test_range_validation
     # exclusive
     assert_raise(ArgumentError) { Axlsx::RangeValidator.validate('foo', 1, 10, 10, false) } 
     # inclusive by default
     assert_nothing_raised { Axlsx::RangeValidator.validate('foo', 1, 10, 10) }
  end
end
