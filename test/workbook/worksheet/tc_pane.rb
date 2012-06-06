# encoding: UTF-8
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
    @symbol_options.each do |key, value|
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
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet :name => "sheetview"
    @ws.sheet_view do |vs|
      vs.pane do |p|
        p.active_pane = :top_left
        p.state = :frozen_split
        p.x_split = 255
        p.y_split = 563
        p.top_left_cell = 'B5'
      end
    end
    
    doc = Nokogiri::XML.parse(@ws.to_xml_string)
    
    assert_equal(1, doc.xpath("//xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView/xmlns:pane[@ySplit='563']
      [@xSplit='255'][@topLeftCell='B5'][@state='frozenSplit'][@activePane='topLeftCell']").size)
    
    assert doc.xpath(  "//xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView/xmlns:pane[@ySplit='563']
      [@xSplit='255'][@topLeftCell='B5'][@state='frozenSplit'][@activePane='topLeftCell']")
  end
end