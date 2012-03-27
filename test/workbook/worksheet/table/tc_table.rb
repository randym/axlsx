require 'tc_helper.rb'

class TestTable < Test::Unit::TestCase
  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet
    40.times do
      @ws << ["aa","aa","aa","aa","aa","aa"]
    end
  end

  def test_initialization
    assert(@ws.workbook.tables.empty?)
    assert(@ws.tables.empty?)

  end

  def test_add_table
    name = "test"
    table = @ws.add_table("A1:D5", :name => name)
    assert(table.is_a?(Axlsx::Table), "must create a table")
    assert_equal(@ws.workbook.tables.last, table, "must be added to workbook table collection")
    assert_equal(@ws.tables.last, table, "must be added to worksheet table collection")
    assert_equal(table.name, name, "options for name are applied")

  end

  def test_charts
    assert(@ws.drawing.charts.empty?)
    chart = @ws.add_chart(Axlsx::Pie3DChart, :title=>"bob", :start_at=>[0,0], :end_at=>[1,1])
    assert_equal(@ws.drawing.charts.last, chart, "add chart is returned")
    chart = @ws.add_chart(Axlsx::Pie3DChart, :title=>"nancy", :start_at=>[1,5], :end_at=>[5,10])
    assert_equal(@ws.drawing.charts.last, chart, "add chart is returned")
  end

  def test_pn
    table = @ws.add_table("A1:D5")
    assert_equal(@ws.tables.first.pn, "tables/table1.xml")
  end

  def test_rId
    table = @ws.add_table("A1:D5")
    assert_equal(@ws.tables.first.rId, "rId1")
  end

  def test_index
    table = @ws.add_table("A1:D5")
    assert_equal(@ws.tables.first.index, @ws.workbook.tables.index(@ws.tables.first))
  end

  def test_relationships
    assert(@ws.relationships.empty?)
    table = @ws.add_table("A1:D5")
    assert_equal(@ws.relationships.size, 1, "adding a table adds a relationship")
    table = @ws.add_table("F1:J5")
    assert_equal(@ws.relationships.size, 2, "adding a table adds a relationship")
  end

  def test_to_xml
    table = @ws.add_table("A1:D5")
    schema = Nokogiri::XML::Schema(File.open(Axlsx::SML_XSD))
    doc = Nokogiri::XML(table.to_xml)
    errors = []
    schema.validate(doc).each do |error|
      errors.push error
      puts error.message
    end
    assert(errors.empty?, "error free validation")
  end

end
