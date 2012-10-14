# encoding: UTF-8
require 'tc_helper.rb'

class TestSelection < Test::Unit::TestCase
  def setup
    @options = { :active_cell => 'A2', :active_cell_id => 1, :pane => :top_left, :sqref => 'A2'  }
    @selection = Axlsx::Selection.new(@options)
  end

  def test_active_cell
    assert_raise(ArgumentError) { @selection.active_cell = :active_cell }
    assert_nothing_raised { @selection.active_cell = "F5" }
    assert_equal(@selection.active_cell, "F5")
  end

  def test_active_cell_id
    assert_raise(ArgumentError) { @selection.active_cell_id = "foo" }
    assert_nothing_raised { @selection.active_cell_id = 11 }
    assert_equal(@selection.active_cell_id, 11)
  end

  def test_pane
    assert_raise(ArgumentError) { @selection.pane = "foo´" }
    assert_nothing_raised { @selection.pane = :bottom_right }
    assert_equal(@selection.pane, "bottomRight")
  end

  def test_sqref
    assert_raise(ArgumentError) { @selection.sqref = :sqref }
    assert_nothing_raised { @selection.sqref = "G32" }
    assert_equal(@selection.sqref, "G32")
  end

  def test_to_xml
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet :name => "sheetview"
    @ws.sheet_view do |vs|
      vs.add_selection(:top_left, { :active_cell => 'B2', :sqref => 'B2' })
      vs.add_selection(:top_right, { :active_cell => 'I10', :sqref => 'I10' })
      vs.add_selection(:bottom_left, { :active_cell => 'E55', :sqref => 'E55' })
      vs.add_selection(:bottom_right, { :active_cell => 'I57', :sqref => 'I57' })
    end

    doc = Nokogiri::XML.parse(@ws.to_xml_string)

    assert_equal(1, doc.xpath("//xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView/xmlns:selection[@sqref='B2'][@pane='topLeft'][@activeCell='B2']").size)
    assert doc.xpath("//xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView/xmlns:selection[@sqref='B2'][@pane='topLeft'][@activeCell='B2']")
    assert_equal(1, doc.xpath("//xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView/xmlns:selection[@sqref='I10'][@pane='topRight'][@activeCell='I10']").size)
    assert doc.xpath("//xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView/xmlns:selection[@sqref='I10'][@pane='topRight'][@activeCell='I10']")
    assert_equal(1, doc.xpath("//xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView/xmlns:selection[@sqref='E55'][@pane='bottomLeft'][@activeCell='E55']").size)
    assert doc.xpath("//xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView/xmlns:selection[@sqref='E55'][@pane='bottomLeft'][@activeCell='E55']")
    assert_equal(1, doc.xpath("//xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView/xmlns:selection[@sqref='I57'][@pane='bottomRight'][@activeCell='I57']").size)
    assert doc.xpath("//xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView/xmlns:selection[@sqref='I57'][@pane='bottomRight'][@activeCell='I57']")
  end
end
