# encoding: UTF-8
$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../../"
require 'tc_helper.rb'

class TestPane < Test::Unit::TestCase
  def setup
    #inverse defaults for booleans
    @nil_options = { :active_pane => :bottom_left, :state => :frozen, :top_left_cell => 'A2' }
    @int_0_options = { :x_split => 2, :y_split => 2 }
    
    @string_options = { :top_left_cell => 'A2' }
    @integer_options = { :x_split => 2, :y_split => 2 }
    @symbol_options = { :active_pane => :bottom_left, :state => :frozen }
    
    @options = @nil_options.merge(@int_0_options)
    
    @pane = Axlsx::Pane.new(@options)
  end
  
  def test_initialize
    pane = Axlsx::Pane.new
    
    @nil_options.each do |key, value|
      assert_equal(nil, pane.send(key.to_sym), "initialized default #{key} should be nil")
      assert_equal(value, @pane.send(key.to_sym), "initialized options #{key} should be #{value}")
    end
    
    @int_0_options.each do |key, value|
      assert_equal(0, pane.send(key.to_sym), "initialized default #{key} should be 0")
      assert_equal(value, @pane.send(key.to_sym), "initialized options #{key} should be #{value}")
    end
  end
  
  def test_string_attribute_validation
    @string_options.each do |key, value|
      assert_raise(ArgumentError, "#{key} must be string") { @pane.send("#{key}=".to_sym, :symbol) }
      assert_nothing_raised { @pane.send("#{key}=".to_sym, "foo") }
    end
  end
  
  def test_symbol_attribute_validation
    @symbol_options.each do |key, value|
      assert_raise(ArgumentError, "#{key} must be symbol") { @pane.send("#{key}=".to_sym, "foo") }
      assert_nothing_raised { @pane.send("#{key}=".to_sym, value) }
    end
  end
  
  def test_integer_attribute_validation
    @integer_options.each do |key, value|
      assert_raise(ArgumentError, "#{key} must be integer") { @pane.send("#{key}=".to_sym, "foo") }
      assert_nothing_raised { @pane.send("#{key}=".to_sym, value) }
    end
  end
  
  def test_active_pane
    assert_raise(ArgumentError) { @pane.active_pane = "10" }
    assert_nothing_raised { @pane.active_pane = :top_left }
    assert_equal(@pane.active_pane, :top_left)
  end
  
  def test_state
    assert_raise(ArgumentError) { @pane.state = "foo" }
    assert_nothing_raised { @pane.state = :frozen_split }
    assert_equal(@pane.state, :frozen_split)
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
