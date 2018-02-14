require 'tc_helper.rb'

class TestSharedStringsTable < Test::Unit::TestCase

  def setup
    @p = Axlsx::Package.new :use_shared_strings=>true

    ws = @p.workbook.add_worksheet
    ws.add_row ['a', 1, 'b']
    ws.add_row ['b', 1, 'c']
    ws.add_row ['c', 1, 'd']
    ws.rows.last.add_cell('b', :type => :text)
  end

  def test_workbook_has_shared_strings
    assert(@p.workbook.shared_strings.is_a?(Axlsx::SharedStringsTable), "shared string table was not created")
  end

  def test_count
    sst = @p.workbook.shared_strings
    assert_equal(sst.count, 7)
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

  def test_remove_control_characters_in_xml_serialization
    nasties =  "hello\x10\x00\x1C\x1Eworld"
    @p.workbook.worksheets[0].add_row [nasties]

    # test that the nasty string was added to the shared strings
    assert @p.workbook.shared_strings.unique_cells.has_key?(nasties)

    # test that none of the control characters are in the XML output for shared strings
    assert_no_match(/#{Axlsx::CONTROL_CHARS}/, @p.workbook.shared_strings.to_xml_string)

    # assert that the shared string was normalized to remove the control characters
    assert_not_nil @p.workbook.shared_strings.to_xml_string.index("helloworld")
  end
end
