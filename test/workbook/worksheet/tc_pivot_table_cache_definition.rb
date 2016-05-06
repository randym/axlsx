require 'tc_helper.rb'

class TestPivotTableCacheDefinition < Test::Unit::TestCase
  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet
    @ws << ['headerString1','headerInt','headerString2','headerString3']
    4.times do |idx|
      @ws << ["value1_#{idx}", idx, "value2_#{idx}", "value3_#{idx}"]
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

  def test_cache_contains_header_values
    data_sheet = @ws.clone
    data_sheet.name = "Pivot Table Data Source"
    @pivot_table.data_sheet = data_sheet
    @pivot_table.rows = ['headerString2']
    doc = Nokogiri::XML(@cache_definition.to_xml_string)
    shared_items_xml = doc.css('cacheFields cacheField')[2].css('sharedItems').first
    assert_equal('4', shared_items_xml['count'])
    assert_equal('1', shared_items_xml['containsSemiMixedTypes'])
    assert_equal('0', shared_items_xml['containsInteger'])
    assert_equal('0', shared_items_xml['containsNumber'])
    assert_equal('1', shared_items_xml['containsString'])
    assert_equal(['value2_0', 'value2_1', 'value2_2', 'value2_3'], shared_items_xml.css('s').map { |xml_item| xml_item['v'] })
  end

  def test_cache_contains_header_values_with_numeric_type
    data_sheet = @ws.clone
    data_sheet.name = "Pivot Table Data Source"
    @pivot_table.data_sheet = data_sheet
    @pivot_table.rows = ['headerInt']
    doc = Nokogiri::XML(@cache_definition.to_xml_string)
    shared_items_xml = doc.css('cacheFields cacheField')[1].css('sharedItems').first
    assert_equal('4', shared_items_xml['count'])
    assert_equal('0', shared_items_xml['containsSemiMixedTypes'])
    assert_equal('1', shared_items_xml['containsInteger'])
    assert_equal('1', shared_items_xml['containsNumber'])
    assert_equal('0', shared_items_xml['containsString'])
    assert_equal('0', shared_items_xml['minValue'])
    assert_equal('3', shared_items_xml['maxValue'])
    assert_equal(['0', '1', '2', '3'], shared_items_xml.css('n').map { |xml_item| xml_item['v'] })
  end

end
