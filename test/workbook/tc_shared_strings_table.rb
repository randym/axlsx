require 'tc_helper.rb'

class TestSharedStringsTable < Test::Unit::TestCase

  def setup
    @p = Axlsx::Package.new :use_shared_strings=>true
    ws = @p.workbook.add_worksheet
    ws.add_row ['a', 1, 'b']
    ws.add_row ['b', 1, 'c']
    ws.add_row ['c', 1, 'd']
  end

  def test_workbook_has_shared_strings
    assert(@p.workbook.shared_strings.is_a?(Axlsx::SharedStringsTable), "shared string table was not created")
  end

  def test_count
    sst = @p.workbook.shared_strings
    assert_equal(sst.count, 6)
  end

  def test_unique_count
    sst = @p.workbook.shared_strings
    assert_equal(sst.unique_count, 4)
  end

  def test_uses_workbook_xml_space
    assert_equal(@p.workbook.xml_space, @p.workbook.shared_strings.xml_space)
    @p.workbook.xml_space = :default
    assert_equal(:default, @p.workbook.shared_strings.xml_space)
  end

  def test_valid_document
    schema = Nokogiri::XML::Schema(File.open(Axlsx::SML_XSD))
    doc = Nokogiri::XML(@p.workbook.shared_strings.to_xml_string)
    errors = []
    schema.validate(doc).each do |error|
      puts error.message
      errors << error
    end
    assert_equal(errors.size, 0, "sharedStirngs.xml Invalid" + errors.map{ |e| e.message }.to_s)
  end

end
