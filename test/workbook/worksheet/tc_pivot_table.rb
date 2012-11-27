require 'tc_helper.rb'

class TestPivotTable < Test::Unit::TestCase
  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet
    40.times do
      @ws << ["aa","aa","aa","aa","aa","aa"]
    end
  end

  def test_initialization
    assert(@ws.workbook.pivot_tables.empty?)
    assert(@ws.pivot_tables.empty?)
  end

  def test_add_pivot_table
    pivot_table = @ws.add_pivot_table('G5:G6', 'A1:D5')
    assert_equal('G5:G6', pivot_table.ref, 'ref assigned from first parameter')
    assert_equal('A1:D5', pivot_table.range, 'range assigned from second parameter')
    assert_equal('PivotTable1', pivot_table.name, 'name automatically generated')
    assert(pivot_table.is_a?(Axlsx::PivotTable), "must create a pivot table")
    assert_equal(@ws.workbook.pivot_tables.last, pivot_table, "must be added to workbook pivot tables collection")
    assert_equal(@ws.pivot_tables.last, pivot_table, "must be added to worksheet pivot tables collection")
  end

  def test_pn
    @ws.add_pivot_table('G5:G6', 'A1:D5')
    assert_equal(@ws.pivot_tables.first.pn, "pivotTables/pivotTable1.xml")
  end

  def test_rId
    @ws.add_pivot_table('G5:G6', 'A1:D5')
    assert_equal(@ws.pivot_tables.first.rId, "rId1")
  end

  def test_index
    @ws.add_pivot_table('G5:G6', 'A1:D5')
    assert_equal(@ws.pivot_tables.first.index, @ws.workbook.pivot_tables.index(@ws.pivot_tables.first))
  end

  def test_relationships
    assert(@ws.relationships.empty?)
    @ws.add_pivot_table('G5:G6', 'A1:D5')
    assert_equal(@ws.relationships.size, 1, "adding a pivot table adds a relationship")
    @ws.add_pivot_table('G10:G11', 'A1:D5')
    assert_equal(@ws.relationships.size, 2, "adding a pivot table adds a relationship")
  end

  def test_to_xml_string
    pivot_table = @ws.add_pivot_table('G5:G6', 'A1:D5')
    schema = Nokogiri::XML::Schema(File.open(Axlsx::SML_XSD))
    doc = Nokogiri::XML(pivot_table.to_xml_string)
    errors = []
    schema.validate(doc).each do |error|
      errors.push error
      puts error.message
    end
    assert(errors.empty?, "error free validation")
  end

end
