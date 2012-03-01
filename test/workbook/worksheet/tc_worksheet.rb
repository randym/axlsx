require 'test/unit'
require 'axlsx.rb'

class TestWorksheet < Test::Unit::TestCase
  def setup    
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet
  end

  def test_pn
    assert_equal(@ws.pn, "worksheets/sheet1.xml")
    ws = @ws.workbook.add_worksheet
    assert_equal(ws.pn, "worksheets/sheet2.xml")
  end

  def test_page_margins
    assert(@ws.page_margins.is_a? Axlsx::PageMargins)    
  end

  def test_page_margins_yeild
    @ws.page_margins do |pm|
      assert(pm.is_a? Axlsx::PageMargins)
      assert(@ws.page_margins == pm)
    end
  end

  def test_initialization_options
    page_margins = {:left => 2, :right => 2, :bottom => 2, :top => 2, :header => 2, :footer => 2}
    optioned = @ws.workbook.add_worksheet(:name => 'bob', :page_margins => page_margins)
    assert_equal(optioned.page_margins.left, page_margins[:left])
    assert_equal(optioned.page_margins.right, page_margins[:right])
    assert_equal(optioned.page_margins.top, page_margins[:top])
    assert_equal(optioned.page_margins.bottom, page_margins[:bottom])
    assert_equal(optioned.page_margins.header, page_margins[:header])
    assert_equal(optioned.page_margins.footer, page_margins[:footer])
    assert_equal(optioned.name, 'bob')
  end

  def test_rels_pn
    assert_equal(@ws.rels_pn, "worksheets/_rels/sheet1.xml.rels")
    ws = @ws.workbook.add_worksheet
    assert_equal(ws.rels_pn, "worksheets/_rels/sheet2.xml.rels")
  end

  def test_rId
    assert_equal(@ws.rId, "rId1")
    ws = @ws.workbook.add_worksheet
    assert_equal(ws.rId, "rId2")
  end
  
  def test_index
    assert_equal(@ws.index, @ws.workbook.worksheets.index(@ws))
  end

  def test_dimension
    @ws.add_row [1, 2, 3]
    @ws.add_row [4, 5, 6]
    assert_equal @ws.dimension, "A1:C2"
  end

  def test_referencing
    @ws.add_row [1, 2, 3]
    @ws.add_row [4, 5, 6]
    range = @ws["A1:C2"]
    first_row = @ws[1]
    last_row = @ws[2]
    assert_equal(@ws.rows[0],first_row)
    assert_equal(@ws.rows[1],last_row)
    assert_equal(range.size, 6)
    assert_equal(range.first, @ws.rows.first.cells.first)
    assert_equal(range.last, @ws.rows.last.cells.last)    
  end

  def test_add_row
    assert(@ws.rows.empty?, "sheet has no rows by default")
    r = @ws.add_row([1,2,3])
    assert_equal(@ws.rows.size, 1, "add_row adds a row")
    assert_equal(@ws.rows.first, r, "the row returned is the row added")
  end

  def test_add_chart
    assert(@ws.workbook.charts.empty?, "the sheet's workbook should not have any charts by default")
    @ws.add_chart Axlsx::Pie3DChart
    assert_equal(@ws.workbook.charts.size, 1, "add_chart adds a chart to the workbook")
  end

  def test_drawing
    assert @ws.drawing.is_a? Axlsx::Drawing
  end

  def test_col_style
    @ws.add_row [1,2,3,4]
    @ws.add_row [1,2,3,4]
    @ws.add_row [1,2,3,4]
    @ws.add_row [1,2,3,4]
    @ws.col_style( (1..2), 1, :row_offset=>1)
    @ws.rows[(1..-1)].each do | r | 
      assert_equal(r.cells[1].style, 1)
      assert_equal(r.cells[2].style, 1)
    end
    assert_equal(@ws.rows.first.cells[1].style, 0)
    assert_equal(@ws.rows.first.cells[0].style, 0)
  end

  def test_col_style_with_empty_column
    @ws.add_row [1,2,3,4]
    @ws.add_row [1]
    @ws.add_row [1,2,3,4]
    assert_nothing_raised {@ws.col_style(1, 1)}
  end

  def test_cols
    @ws.add_row [1,2,3,4]
    @ws.add_row [1,2,3,4]
    @ws.add_row [1,2,3,4]
    @ws.add_row [1,2,3,4]
    c = @ws.cols[1]
    assert_equal(c.size, 4)    
    assert_equal(c[0].value, 2)
  end

  def test_row_style
    @ws.add_row [1,2,3,4]
    @ws.add_row [1,2,3,4]
    @ws.add_row [1,2,3,4]
    @ws.add_row [1,2,3,4]
    @ws.row_style 1, 1, :col_offset=>1
    @ws.rows[1].cells[(1..-1)].each do | c | 
      assert_equal(c.style, 1)
    end
    assert_equal(@ws.rows[1].cells[0].style, 0)
    assert_equal(@ws.rows[2].cells[1].style, 0)
  end
  
  def test_to_xml
    schema = Nokogiri::XML::Schema(File.open(Axlsx::SML_XSD))
    doc = Nokogiri::XML(@ws.to_xml)
    errors = []
    schema.validate(doc).each do |error|
      errors.push error
      puts error.message
    end
    assert(errors.empty?, "error free validation")
  end

  def test_valid_with_page_margins
    @ws.page_margins.set :left => 9
    schema = Nokogiri::XML::Schema(File.open(Axlsx::SML_XSD))
    doc = Nokogiri::XML(@ws.to_xml)
    errors = []
    schema.validate(doc).each do |error|
      errors.push error
      puts error.message
    end
    assert(errors.empty?, "error free validation")
    
  end

  def test_relationships
    assert(@ws.relationships.empty?, "No Drawing relationship until you add a chart")
    c = @ws.add_chart Axlsx::Pie3DChart
    assert_equal(@ws.relationships.size, 1, "adding a chart creates the relationship")
    c = @ws.add_chart Axlsx::Pie3DChart
    assert_equal(@ws.relationships.size, 1, "multiple charts still only result in one relationship")
  end

  
  def test_name_unique
    assert_raise(ArgumentError, "worksheet name must be unique") { n = @ws.name; @ws.workbook.add_worksheet(:name=> @ws) }
  end

  def test_name_size
    assert_raise(ArgumentError, "name too long!") { @ws.name = Array.new(32, "A").join('') }
    assert_nothing_raised { @ws.name = Array.new(31, "A").join('') }
  end

  def test_update_auto_with_data
    small = @ws.workbook.styles.add_style(:sz=>2)
    big = @ws.workbook.styles.add_style(:sz=>10)

    @ws.add_row ["chasing windmills", "penut"], :style=>small
    assert(@ws.auto_fit_data.size == 2, "a data item for each column")

    assert_equal(@ws.auto_fit_data[0], {:sz => 2, :longest => "chasing windmills", :fixed=>nil}, "adding a row updates auto_fit_data if the product of the string length and font is greater for the column")


    @ws.add_row ["mule"], :style=>big
    assert_equal(@ws.auto_fit_data[0], {:sz=>10,:longest=>"mule", :fixed=>nil}, "adding a row updates auto_fit_data if the product of the string length and font is greater for the column")
  end

  def test_set_fixed_width_column
    @ws.add_row ["mule", "donkey", "horse"], :widths => [20, :ignore, nil]
    assert(@ws.auto_fit_data.size == 3, "a data item for each column")
    assert_equal({:sz=>11, :longest=>"mule", :fixed=>20 }, @ws.auto_fit_data[0], "adding a row with fixed width updates :fixed attribute")
    assert_equal({:sz=>11, :longest=>"", :fixed=>nil}, @ws.auto_fit_data[1], ":ignore does not set any data")
    assert_equal({:sz=>11, :longest=>"horse", :fixed=>nil}, @ws.auto_fit_data[2], "nil, well really anything else just works as normal")
    @ws.add_row ["mule", "donkey", "horse"]
    assert_equal({:sz=>11, :longest=>"donkey", :fixed=>nil}, @ws.auto_fit_data[1])

  end

  def test_fixed_widths_with_merged_cells
    @ws.add_row ["hey, I'm like really long and stuff so I think you will merge me."]
    @ws.merge_cells "A1:C1"
    @ws.add_row ["but Im Short!"], :widths=> [14.8]
    assert_equal(@ws.send(:auto_width, @ws.auto_fit_data[0]), 14.8)
  end
  
  def test_fixed_width_to_auto
    @ws.add_row ["hey, I'm like really long and stuff so I think you will merge me."]
    @ws.merge_cells "A1:C1"
    @ws.add_row ["but Im Short!"], :widths=> [14.8]
    assert_equal(@ws.send(:auto_width, @ws.auto_fit_data[0]), 14.8)
    @ws.add_row ["no, I like auto!"], :widths=>[:auto]
    assert_equal(@ws.auto_fit_data[0][:fixed], nil)
  end

  def test_auto_width
    assert(@ws.send(:auto_width, {:sz=>11, :longest=>"fisheries"}) > @ws.send(:auto_width, {:sz=>11, :longest=>"fish"}), "longer strings get a longer auto_width at the same font size")

    assert(@ws.send(:auto_width, {:sz=>11, :longest=>"fish"}) < @ws.send(:auto_width, {:sz=>12, :longest=>"fish"}), "larger fonts produce longer with with same string")
    assert_equal(@ws.send(:auto_width, {:sz=>11, :longest => "This is a really long string", :fixed=>0.2}), 0.2, "fixed rules!")
  end

  def test_fixed_height
    @ws.add_row [1, 2, 3], :height => 40
    assert_equal(40, @ws.rows[-1].height)
  end


  def test_set_column_width
    @ws.add_row ["chasing windmills", "penut"]
    assert_equal(@ws.auto_fit_data[0][:fixed], nil, 'no fixed by default')
    @ws.column_widths nil, 0.5
    assert_equal(@ws.auto_fit_data[1][:fixed], 0.5, 'eat my width')
    assert_raise(ArgumentError, 'reject invalid columns') { @ws.column_widths 2, 7, nil }
    assert_raise(ArgumentError, 'only accept unsigned ints') { @ws.column_widths 2, 7, -1 }
    assert_raise(ArgumentError, 'only accept Integer, Float or Fixnum') { @ws.column_widths 2, 7, "-1" }
  end

  def test_merge_cells
    assert(@ws.merged_cells.is_a?(Array))
    assert_equal(@ws.merged_cells.size, 0)           
    @ws.add_row [1,2,3]
    @ws.add_row [4,5,6]
    @ws.add_row [7,8,9]
    @ws.merge_cells "A1:A2"
    @ws.merge_cells "B2:C3"
    @ws.merge_cells @ws.rows.last.cells[(0..1)]
    assert_equal(@ws.merged_cells.size, 3)
    assert_equal(@ws.merged_cells.last, "A3:B3")
  end

  def test_auto_filter
    assert(@ws.auto_filter.nil?)
    assert_raise(ArgumentError) { @ws.auto_filter = 123 }
    @ws.auto_filter = "A1:D9"
    assert_equal(@ws.auto_filter, "A1:D9")
  end
end
