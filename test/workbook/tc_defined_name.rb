require 'tc_helper'

class TestDefinedNames < Minitest::Unit::TestCase
  def setup
    @dn = Axlsx::DefinedName.new('Sheet1!A1:A1')
  end

  def test_initialize
    assert_equal('Sheet1!A1:A1', @dn.formula)
  end

  def test_string_attributes
    %w(short_cut_key status_bar help description custom_menu comment).each do |attr|
      assert_raises(ArgumentError, 'only strings allowed in string attributes') { @dn.send("#{attr}=", 1) }
      assert_nothing_raised { @dn.send("#{attr}=", '_xlnm.Sheet_Title') }
    end
  end

  def test_boolean_attributes
   %w(workbook_parameter publish_to_server xlm vb_proceedure function hidden).each do |attr|
      assert_raises(ArgumentError, 'only booleanish allowed in string attributes') { @dn.send("#{attr}=", 'foo') }
      assert_nothing_raised { @dn.send("#{attr}=", 1) }
    end

  end

  def test_local_sheet_id
    assert_raises(ArgumentError, 'local_sheet_id must be an unsigned int') { @dn.local_sheet_id = -1 }
    assert_nothing_raised { @dn.local_sheet_id = 1 }
  end

  def test_do_not_camelcase_value_for_name
    @dn.name = '_xlnm._FilterDatabase'
    doc = Nokogiri::XML(@dn.to_xml_string)
    assert_equal(doc.xpath("//definedName[@name='_xlnm._FilterDatabase']").size, 1)
    assert_equal('Sheet1!A1:A1', doc.xpath('//definedName').text)
  end

  def test_to_xml_string
    assert_raises(ArgumentError, 'name is required for serialization') { @dn.to_xml_string }
    @dn.name = '_xlnm.Print_Titles'
    @dn.hidden = true
    doc = Nokogiri::XML(@dn.to_xml_string)
    assert_equal(doc.xpath("//definedName[@name='_xlnm.Print_Titles']").size, 1)
    assert_equal(doc.xpath("//definedName[@hidden='1']").size, 1)
    assert_equal('Sheet1!A1:A1', doc.xpath('//definedName').text)
  end

end
