# encoding: UTF-8
require 'tc_helper.rb'

# <xsd:complexType name="CT_SheetProtection">
# <xsd:attribute name="algorithmName" type="s:ST_Xstring" use="optional"/>
# <xsd:attribute name="hashValue" type="xsd:base64Binary" use="optional"/>
# <xsd:attribute name="saltValue" type="xsd:base64Binary" use="optional"/>
# <xsd:attribute name="spinCount" type="xsd:unsignedInt" use="optional"/>
# <xsd:attribute name="sheet" type="xsd:boolean" use="optional" default="false"/>
# <xsd:attribute name="objects" type="xsd:boolean" use="optional" default="false"/>
# <xsd:attribute name="scenarios" type="xsd:boolean" use="optional" default="false"/>
# <xsd:attribute name="formatCells" type="xsd:boolean" use="optional" default="true"/>
# <xsd:attribute name="formatColumns" type="xsd:boolean" use="optional" default="true"/>
# <xsd:attribute name="formatRows" type="xsd:boolean" use="optional" default="true"/>
# <xsd:attribute name="insertColumns" type="xsd:boolean" use="optional" default="true"/>
# <xsd:attribute name="insertRows" type="xsd:boolean" use="optional" default="true"/>
# <xsd:attribute name="insertHyperlinks" type="xsd:boolean" use="optional" default="true"/>
# <xsd:attribute name="deleteColumns" type="xsd:boolean" use="optional" default="true"/>
# <xsd:attribute name="deleteRows" type="xsd:boolean" use="optional" default="true"/>
# <xsd:attribute name="selectLockedCells" type="xsd:boolean" use="optional" default="false"/>
# <xsd:attribute name="sort" type="xsd:boolean" use="optional" default="true"/>
# <xsd:attribute name="autoFilter" type="xsd:boolean" use="optional" default="true"/>
# <xsd:attribute name="pivotTables" type="xsd:boolean" use="optional" default="true"/>
# <xsd:attribute name="selectUnlockedCells" type="xsd:boolean" use="optional" default="false"/>
# </xsd:complexType>
 
class TestSheetProtection < Test::Unit::TestCase
  def setup
    #inverse defaults
    @options = { :sheet => true, :objects => true, :scenarios => true, :format_cells => false,
                :format_columns => false, :format_rows => false, :insert_columns => false, :insert_rows => false,
                :insert_hyperlinks => false, :delete_columns => false, :delete_rows => false, :select_locked_cells => true,
                :sort => false, :auto_filter => false, :pivot_tables => false, :select_unlocked_cells => true }
    @sp = Axlsx::SheetProtection.new(@options)
  end

  def test_initialize
    sp = Axlsx::SheetProtection.new
    @options.each do |key, value|
      assert_equal(!value, sp.send(key.to_sym), "initialized default #{key} should be #{!value}")
      assert_equal(value, @sp.send(key.to_sym), "initialized options #{key} should be #{value}")
    end
    { :algorithm_name => nil }.each do |key, value|
      assert_equal(value, @sp.send(key.to_sym), "initialized default #{key} should be #{value}")
      assert_equal(value, sp.send(key.to_sym), "initialized default #{key} should be #{value}")
    end 
  end

  def test_boolean_attribute_validation
    @options.each do |key, value|
      assert_raise(ArgumentError, "#{key} must be boolean") { @sp.send("#{key}=".to_sym, 'A') }
      assert_nothing_raised { @sp.send("#{key}=".to_sym, true) }
      assert_nothing_raised { @sp.send("#{key}=".to_sym, true) }
    end
  end
  
 
  def test_to_xml_string
    @sp.password = 'fish'
    doc = Nokogiri::XML(@sp.to_xml_string)
    @options.each do |key, value|
      assert(doc.xpath("//sheetProtection[@#{key.to_s.gsub(/_(.)/){ $1.upcase }}='#{value}']")) 
    end
  end

end






















































