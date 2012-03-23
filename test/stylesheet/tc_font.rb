require 'tc_helper.rb'

class TestFont < Test::Unit::TestCase

  def setup
    @item = Axlsx::Font.new
  end

  def teardown
  end


  def test_initialiation
    assert_equal(@item.name, nil)
    assert_equal(@item.charset, nil)
    assert_equal(@item.family, nil)
    assert_equal(@item.b, nil)
    assert_equal(@item.i, nil)
    assert_equal(@item.u, nil)
    assert_equal(@item.strike, nil)
    assert_equal(@item.outline, nil)
    assert_equal(@item.shadow, nil)
    assert_equal(@item.condense, nil)
    assert_equal(@item.extend, nil)
    assert_equal(@item.color, nil)
    assert_equal(@item.sz, nil)
  end



    # def name=(v) Axlsx::validate_string v; @name = v end
  def test_name
    assert_raise(ArgumentError) { @item.name = 7 }
    assert_nothing_raised { @item.name = "bob" }
    assert_equal(@item.name, "bob")
  end
    # def charset=(v) Axlsx::validate_unsigned_int v; @charset = v end
  def test_charset
    assert_raise(ArgumentError) { @item.charset = -7 }
    assert_nothing_raised { @item.charset = 5 }
    assert_equal(@item.charset, 5)
  end

    # def family=(v) Axlsx::validate_unsigned_int v; @family = v end
  def test_family
    assert_raise(ArgumentError) { @item.family = -7 }
    assert_nothing_raised { @item.family = 5 }
    assert_equal(@item.family, 5)
  end

    # def b=(v) Axlsx::validate_boolean v; @b = v end
  def test_b
    assert_raise(ArgumentError) { @item.b = -7 }
    assert_nothing_raised { @item.b = true }
    assert_equal(@item.b, true)
  end

    # def i=(v) Axlsx::validate_boolean v; @i = v end
  def test_i
    assert_raise(ArgumentError) { @item.i = -7 }
    assert_nothing_raised { @item.i = true }
    assert_equal(@item.i, true)
  end

    # def u=(v) Axlsx::validate_boolean v; @u = v end
  def test_u
    assert_raise(ArgumentError) { @item.u = -7 }
    assert_nothing_raised { @item.u = true }
    assert_equal(@item.u, true)
  end

    # def strike=(v) Axlsx::validate_boolean v; @strike = v end
  def test_strike
    assert_raise(ArgumentError) { @item.strike = -7 }
    assert_nothing_raised { @item.strike = true }
    assert_equal(@item.strike, true)
  end

    # def outline=(v) Axlsx::validate_boolean v; @outline = v end
  def test_outline
    assert_raise(ArgumentError) { @item.outline = -7 }
    assert_nothing_raised { @item.outline = true }
    assert_equal(@item.outline, true)
  end

    # def shadow=(v) Axlsx::validate_boolean v; @shadow = v end
  def test_shadow
    assert_raise(ArgumentError) { @item.shadow = -7 }
    assert_nothing_raised { @item.shadow = true }
    assert_equal(@item.shadow, true)
  end

    # def condense=(v) Axlsx::validate_boolean v; @condense = v end
  def test_condense
    assert_raise(ArgumentError) { @item.condense = -7 }
    assert_nothing_raised { @item.condense = true }
    assert_equal(@item.condense, true)
  end

    # def extend=(v) Axlsx::validate_boolean v; @extend = v end
  def test_extend
    assert_raise(ArgumentError) { @item.extend = -7 }
    assert_nothing_raised { @item.extend = true }
    assert_equal(@item.extend, true)
  end

    # def color=(v) DataTypeValidator.validate "Font.color", Color, v; @color=v end
  def test_color
    assert_raise(ArgumentError) { @item.color = -7 }
    assert_nothing_raised { @item.color = Axlsx::Color.new(:rgb=>"00000000") }
    assert(@item.color.is_a?(Axlsx::Color))
  end

    # def sz=(v) Axlsx::validate_unsigned_int v; @sz=v end
  def test_sz
    assert_raise(ArgumentError) { @item.sz = -7 }
    assert_nothing_raised { @item.sz = 5 }
    assert_equal(@item.sz, 5)
  end

end
