require 'tc_helper.rb'

class TestView3D < Minitest::Unit::TestCase
  def setup
    @view  = Axlsx::View3D.new
  end

  def teardown
  end

  def test_options
    v = Axlsx::View3D.new :rot_x => 10, :rot_y => 5, :h_percent => "30%", :depth_percent => "45%", :r_ang_ax => false, :perspective => 10
    assert_equal(v.rot_x, 10)
    assert_equal(v.rot_y, 5)
    assert_equal(v.h_percent, "30%")
    assert_equal(v.depth_percent, "45%")
    assert_equal(v.r_ang_ax, false)
    assert_equal(v.perspective, 10)
  end

  def test_rot_x
    assert_raises(ArgumentError) {@view.rot_x = "bob"}
    assert_nothing_raised {@view.rot_x = -90}
  end

  def test_rot_y
    assert_raises(ArgumentError) {@view.rot_y = "bob"}
    assert_nothing_raised {@view.rot_y = 90}
  end

  def test_h_percent
    assert_raises(ArgumentError) {@view.h_percent = "bob"}
    assert_nothing_raised {@view.h_percent = "500%"}
  end

  def test_depth_percent
    assert_raises(ArgumentError) {@view.depth_percent = "bob"}
    assert_nothing_raised {@view.depth_percent = "20%"}
  end


  def test_rAngAx
    assert_raises(ArgumentError) {@view.rAngAx = "bob"}
    assert_nothing_raised {@view.rAngAx = true}
  end

  def test_perspective
    assert_raises(ArgumentError) {@view.perspective = "bob"}
    assert_nothing_raised {@view.perspective = 30}
  end



end
