require 'test/unit'
require 'axlsx.rb'

class TestView3D < Test::Unit::TestCase
  def setup    
    @view  = Axlsx::View3D.new
  end

  def teardown
  end
  
  def test_options
    v = Axlsx::View3D.new :rotX => 10, :rotY => 5, :hPercent => "30%", :depthPercent => "45%", :rAngAx => false, :perspective => 10
    assert_equal(v.rotX, 10)
    assert_equal(v.rotY, 5)
    assert_equal(v.hPercent, "30%")
    assert_equal(v.depthPercent, "45%")
    assert_equal(v.rAngAx, false)
    assert_equal(v.perspective, 10)
  end
  
  def test_rotX
    assert_raise(ArgumentError) {@view.rotX = "bob"}
    assert_nothing_raised {@view.rotX = -90}
  end

  def test_rotY
    assert_raise(ArgumentError) {@view.rotY = "bob"}
    assert_nothing_raised {@view.rotY = 90}
  end

  def test_hPercent
    assert_raise(ArgumentError) {@view.hPercent = "bob"}
    assert_nothing_raised {@view.hPercent = "500%"}
  end

  def test_depthPercent
    assert_raise(ArgumentError) {@view.depthPercent = "bob"}
    assert_nothing_raised {@view.depthPercent = "20%"}
  end


  def test_rAngAx
    assert_raise(ArgumentError) {@view.rAngAx = "bob"}
    assert_nothing_raised {@view.rAngAx = true}
  end
  
  def test_perspective
    assert_raise(ArgumentError) {@view.perspective = "bob"}
    assert_nothing_raised {@view.perspective = 30}
  end


  
end
