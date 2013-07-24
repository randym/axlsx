require 'tc_helper.rb'

class TestPivotTableCacheDefinition < Test::Unit::TestCase
  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet
    5.times do
      @ws << ["aa","aa","aa","aa"]
    end
    @pivot_table = @ws.add_pivot_table('G5:G6', 'A1:D5')
    @cache_definition = @pivot_table.cache_definition
  end

  def test_initialization
    assert(@cache_definition.is_a?(Axlsx::PivotTableCacheDefinition), "must create a pivot table cache definition")
    assert_equal(@pivot_table, @cache_definition.pivot_table, 'refers back to its pivot table')
  end

  def test_pn
    assert_equal('pivotCache/pivotCacheDefinition1.xml', @cache_definition.pn)
  end

  def test_rId
    assert_equal @pivot_table.relationships.for(@cache_definition).Id, @cache_definition.rId
  end

  def test_index
    assert_equal(0, @cache_definition.index)
  end

  def test_cache_id
    assert_equal(1, @cache_definition.cache_id)
  end

  def test_data_sheet
    data_sheet = @ws.clone
    data_sheet.name = "Pivot Table Data Source"
    @pivot_table.data_sheet = data_sheet

    assert(@cache_definition.to_xml_string.include?(data_sheet.name), "must set the data source correctly")
  end

  def test_to_xml_string
    schema = Nokogiri::XML::Schema(File.open(Axlsx::SML_XSD))
    doc = Nokogiri::XML(@cache_definition.to_xml_string)
    errors = []
    schema.validate(doc).each do |error|
      errors.push error
      puts error.message
    end
    assert(errors.empty?, "error free validation")
  end

end
