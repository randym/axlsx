# encoding: UTF-8
$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../../"
require 'tc_helper.rb'

class TestPane < Test::Unit::TestCase
  def setup
    #inverse defaults for booleans
    @nil_options = { :active_pane => :bottom_left, :state => :frozen, :top_left_cell => 'A2' }
    @int_0_options = { :x_split => 2, :y_split => 2 }
   @options = @nil_options.merge(@int_0_options)
    @pane = Axlsx::Pane.new(@options)
  end


  def test_active_pane
    assert_raise(ArgumentError) { @pane.active_pane = "10" }
    assert_nothing_raised { @pane.active_pane = :top_left }
    assert_equal(@pane.active_pane, "topLeft")
  end
  
  def test_state
    assert_raise(ArgumentError) { @pane.state = "foo" }
    assert_nothing_raised { @pane.state = :frozen_split }
    assert_equal(@pane.state, "frozenSplit")
  end
  
  def test_x_split
    assert_raise(ArgumentError) { @pane.x_split = "foo´" }
    assert_nothing_raised { @pane.x_split = 200 }
    assert_equal(@pane.x_split, 200)
  end
  
  def test_y_split
    assert_raise(ArgumentError) { @pane.y_split = 'foo' }
    assert_nothing_raised { @pane.y_split = 300 }
    assert_equal(@pane.y_split, 300)
  end
  
  def test_top_left_cell
    assert_raise(ArgumentError) { @pane.top_left_cell = :cell }
    assert_nothing_raised { @pane.top_left_cell = "A2" }
    assert_equal(@pane.top_left_cell, "A2")
  end
  
  def test_to_xml
    doc = Nokogiri::XML.parse(@pane.to_xml_string)
    assert_equal(1, doc.xpath("//pane[@ySplit=2][@xSplit='2'][@topLeftCell='A2'][@state='frozen'][@activePane='bottomLeft']").size)
  end
  def test_to_xml_frozen
    pane = Axlsx::Pane.new :state => :frozen, :y_split => 2
    doc = Nokogiri::XML(pane.to_xml_string)
    assert_equal(1, doc.xpath("//pane[@topLeftCell='A3']").size)
  end
end
