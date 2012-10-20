require 'tc_helper.rb'

class TestWorkbook < Test::Unit::TestCase
  def setup
    p = Axlsx::Package.new
    @wb = p.workbook
  end

  def teardown
  end

  def test_no_autowidth
    assert_equal(@wb.use_autowidth, true)
    assert_raise(ArgumentError) {@wb.use_autowidth = 0.1}
    assert_nothing_raised {@wb.use_autowidth = false}
    assert_equal(@wb.use_autowidth, false)
  end


  def test_sheet_by_name_retrieval
    @wb.add_worksheet(:name=>'foo')
    @wb.add_worksheet(:name=>'bar')
    assert_equal('foo', @wb.sheet_by_name('foo').name)
    
  end
  def test_date1904
    assert_equal(Axlsx::Workbook.date1904, @wb.date1904)
    @wb.date1904 = :false
    assert_equal(Axlsx::Workbook.date1904, @wb.date1904)
    Axlsx::Workbook.date1904 = :true
    assert_equal(Axlsx::Workbook.date1904, @wb.date1904)
  end

  def test_add_defined_name
    @wb.add_defined_name 'Sheet1!1:1', :name => '_xlnm.Print_Titles', :hidden => true
    assert_equal(1, @wb.defined_names.size)
  end

  def test_shared_strings
    assert_equal(@wb.use_shared_strings, nil)
    assert_raise(ArgumentError) {@wb.use_shared_strings = 'bpb'}
    assert_nothing_raised {@wb.use_shared_strings = :true}
  end

  def test_add_worksheet
    assert(@wb.worksheets.empty?, "worbook has no worksheets by default")
    ws = @wb.add_worksheet(:name=>"bob")
    assert_equal(@wb.worksheets.size, 1, "add_worksheet adds a worksheet!")
    assert_equal(@wb.worksheets.first, ws, "the worksheet returned is the worksheet added")
    assert_equal(ws.name, "bob", "name option gets passed to worksheet")
  end
  
  def test_insert_worksheet
    @wb.add_worksheet(:name => 'A')
    @wb.add_worksheet(:name => 'B')
    ws3 = @wb.insert_worksheet(0, :name => 'C')
    assert_equal(ws3.name, @wb.worksheets.first.name)
  end

  def test_relationships
    #current relationship size is 1 due to style relation
    assert(@wb.relationships.size == 1)
    @wb.add_worksheet
    assert(@wb.relationships.size == 2)
    @wb.use_shared_strings = true
    assert(@wb.relationships.size == 3)
  end

  def test_to_xml
    schema = Nokogiri::XML::Schema(File.open(Axlsx::SML_XSD))
    doc = Nokogiri::XML(@wb.to_xml_string)
    errors = []
    schema.validate(doc).each do |error|
      errors.push error
      puts error.message
    end
    assert(errors.empty?, "error free validation")
  end
  def test_range_requires__valid_sheet
    ws = @wb.add_worksheet :name=>'fish'
    ws.add_row [1,2,3]
    ws.add_row [4,5,6]
    assert_raise(ArgumentError, "no sheet name part") { @wb["A1:C2"]}
    assert_equal @wb['fish!A1:C2'].size, 6
  end

  def test_to_xml_adds_worksheet_when_worksheets_is_empty
    assert(@wb.worksheets.empty?)
    @wb.to_xml_string
    assert(@wb.worksheets.size == 1)
  end

  def test_to_xml_string_defined_names
    @wb.add_worksheet do |sheet|
      sheet.add_row [1, "two"]
      sheet.auto_filter = "A1:B1"
    end
    doc = Nokogiri::XML(@wb.to_xml_string)
    assert_equal(doc.xpath('//xmlns:workbook/xmlns:definedNames/xmlns:definedName').inner_text, @wb.worksheets[0].auto_filter.defined_name)
  end


end
