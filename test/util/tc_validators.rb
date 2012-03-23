require 'tc_helper.rb'
class TestValidators < Test::Unit::TestCase
  def setup
  end
  def teardown
  end

  def test_validators
    #unsigned_int
    assert_nothing_raised { Axlsx.validate_unsigned_int 1 }
    assert_nothing_raised { Axlsx.validate_unsigned_int +1 }
    assert_raise(ArgumentError) { Axlsx.validate_unsigned_int -1 }
    assert_raise(ArgumentError) { Axlsx.validate_unsigned_int '1' }

    #int
    assert_nothing_raised { Axlsx.validate_int 1 }
    assert_nothing_raised { Axlsx.validate_int -1 }
    assert_raise(ArgumentError) { Axlsx.validate_int 'a' }
    assert_raise(ArgumentError) { Axlsx.validate_int Array }

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

  end
end

