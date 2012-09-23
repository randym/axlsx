require 'tc_helper.rb'

class TestTableStyleInfo < Test::Unit::TestCase
  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet
    40.times do 
      @ws.add_row %w(aa bb cc dd ee ff gg hh ii jj kk)
    end
    @table = @ws.add_table(Axlsx::cell_range([@ws.rows.first.cells.first,@ws.rows.last.cells.last], false), :name => 'foo')
    @options =  { :show_first_column => 1,
                :show_last_column => 1,
                :show_row_stripes => 1,
                :show_column_stripes => 1,
                :name => "TableStyleDark4" }


  end

  def test_initialize
    table_style = Axlsx::TableStyleInfo.new @options
    @options.each do |key, value|
      assert_equal(value, table_style.send(key.to_sym))
    end
  end

  def test_boolean_properties
    table_style = Axlsx::TableStyleInfo.new
    @options.keys.each do |key|
     assert_nothing_raised { table_style.send("#{key.to_sym}=", true) }
     assert_raises(ArgumentError) { table_style.send(key.to_sym, 'foo') }
    end
  end
  def doc
    @doc ||= Nokogiri::XML(Axlsx::TableStyleInfo.new(@options).to_xml_string)
  end

  def test_to_xml_string_first_column
    assert(doc.xpath('//tableStyleInfo[@showLastColumn=1]'))
  end

  def test_to_xml_string_row_stripes
    assert(doc.xpath('//tableStyleInfo[@showRowStripes=1]'))
  end

  def test_to_xml_string_column_stripes
    assert(doc.xpath('//tableStyleInfo[@showColumnStripes=1]'))
  end

  def test_to_xml_string_name
    assert(doc.xpath("//tableStyleInfo[@name=#{@options[:name]}]"))
  end
end
