require 'tc_helper.rb'


def shared_test_pivot_table_xml_validity(pivot_table)
  schema = Nokogiri::XML::Schema(File.open(Axlsx::SML_XSD))
  doc = Nokogiri::XML(pivot_table.to_xml_string)
  errors = []
  schema.validate(doc).each do |error|
    errors.push error
    puts error.message
  end
  assert(errors.empty?, "error free validation")
end

class TestPivotTable < Test::Unit::TestCase
  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet

    @ws << ["Year","Month","Region", "Type", "Sales"]
    @ws << [2012,  "Nov",  "East",   "Soda", "12345"]
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

  def test_set_pivot_table_data_sheet
    pivot_table = @ws.add_pivot_table('G5:G6', 'A1:D5')
    data_sheet = @ws.clone
    data_sheet.name = "Pivot Table Data Source"

    assert_equal(pivot_table.data_sheet.name, @ws.name, "must default to the same sheet the pivot table is added to")
    pivot_table.data_sheet = data_sheet
    assert_equal(pivot_table.data_sheet.name, data_sheet.name, "data sheet assigned to pivot table")
  end

  def test_add_pivot_table_with_config
    pivot_table = @ws.add_pivot_table('G5:G6', 'A1:E5') do |pt|
      pt.rows = ['Year', 'Month']
      pt.columns = ['Type']
      pt.data = ['Sales']
      pt.pages = ['Region']
    end
    assert_equal(['Year', 'Month'], pivot_table.rows)
    assert_equal(['Type'], pivot_table.columns)
    assert_equal([{:ref=>"Sales"}], pivot_table.data)
    assert_equal(['Region'], pivot_table.pages)
    shared_test_pivot_table_xml_validity(pivot_table)
  end

  def test_add_pivot_table_with_options_on_data_field
    pivot_table = @ws.add_pivot_table('G5:G6', 'A1:D5') do |pt|
      pt.data = [{:ref=>"Sales", :subtotal => 'average'}]
    end
    assert_equal([{:ref=>"Sales", :subtotal => 'average'}], pivot_table.data)
  end

  def test_header_indices
    pivot_table = @ws.add_pivot_table('G5:G6', 'A1:E5')
    assert_equal(0,   pivot_table.header_index_of('Year'   ))
    assert_equal(1,   pivot_table.header_index_of('Month'  ))
    assert_equal(2,   pivot_table.header_index_of('Region' ))
    assert_equal(3,   pivot_table.header_index_of('Type'   ))
    assert_equal(4,   pivot_table.header_index_of('Sales'  ))
    assert_equal(nil, pivot_table.header_index_of('Missing'))
    assert_equal(%w(A1 B1 C1 D1 E1), pivot_table.header_cell_refs)
  end

  def test_pn
    @ws.add_pivot_table('G5:G6', 'A1:D5')
    assert_equal(@ws.pivot_tables.first.pn, "pivotTables/pivotTable1.xml")
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
    shared_test_pivot_table_xml_validity(pivot_table)
  end

  def test_to_xml_string_with_configuration
    pivot_table = @ws.add_pivot_table('G5:G6', 'A1:E5') do |pt|
      pt.rows = ['Year', 'Month']
      pt.columns = ['Type']
      pt.data = ['Sales']
      pt.pages = ['Region']
    end
    shared_test_pivot_table_xml_validity(pivot_table)
  end

  def test_to_xml_string_with_options_on_data_field
    pivot_table = @ws.add_pivot_table('G5:G6', 'A1:E5') do |pt|
      pt.data = [{:ref=>"Sales", :subtotal => 'average'}]
    end
    shared_test_pivot_table_xml_validity(pivot_table)
  end
end
