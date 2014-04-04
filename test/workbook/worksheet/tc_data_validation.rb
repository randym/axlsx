# encoding: UTF-8
require 'tc_helper.rb'


class TestDataValidation < Test::Unit::TestCase
  def setup
    #inverse defaults
    @boolean_options = { :allowBlank => false, :showDropDown => true, :showErrorMessage => false, :showInputMessage => true }
    @nil_options = { :formula1 => 'foo',  :formula2 => 'foo', :errorTitle => 'foo', :operator => :lessThan, :prompt => 'foo', :promptTitle => 'foo', :sqref => 'foo' }
    @type_option = { :type => :whole }
    @error_style_option = { :errorStyle => :warning }
    
    @string_options = { :formula1 => 'foo', :formula2 => 'foo', :error => 'foo', :errorTitle => 'foo', :prompt => 'foo', :promptTitle => 'foo', :sqref => 'foo' }
    @symbol_options = { :errorStyle => :warning, :operator => :lessThan, :type => :whole}
    
    @options = @boolean_options.merge(@nil_options).merge(@type_option).merge(@error_style_option)
    
    @dv = Axlsx::DataValidation.new(@options)
  end
  
  def test_initialize
    dv = Axlsx::DataValidation.new
    
    @boolean_options.each do |key, value|
      assert_equal(!value, dv.send(key.to_sym), "initialized default #{key} should be #{!value}")
      assert_equal(value, @dv.send(key.to_sym), "initialized options #{key} should be #{value}")
    end
    
    @nil_options.each do |key, value|
      assert_equal(nil, dv.send(key.to_sym), "initialized default #{key} should be nil")
      assert_equal(value, @dv.send(key.to_sym), "initialized options #{key} should be #{value}")
    end
    
    @type_option.each do |key, value|
      assert_equal(:none, dv.send(key.to_sym), "initialized default #{key} should be :none")
      assert_equal(value, @dv.send(key.to_sym), "initialized options #{key} should be #{value}")
    end
    
    @error_style_option.each do |key, value|
      assert_equal(:stop, dv.send(key.to_sym), "initialized default #{key} should be :stop")
      assert_equal(value, @dv.send(key.to_sym), "initialized options #{key} should be #{value}")
    end
  end
  
  def test_boolean_attribute_validation
    @boolean_options.each do |key, value|
      assert_raise(ArgumentError, "#{key} must be boolean") { @dv.send("#{key}=".to_sym, 'A') }
      assert_nothing_raised { @dv.send("#{key}=".to_sym, true) }
    end
  end
  
  def test_string_attribute_validation
    @string_options.each do |key, value|
      assert_raise(ArgumentError, "#{key} must be string") { @dv.send("#{key}=".to_sym, :symbol) }
      assert_nothing_raised { @dv.send("#{key}=".to_sym, "foo") }
    end
  end
  
  def test_symbol_attribute_validation
    @symbol_options.each do |key, value|
      assert_raise(ArgumentError, "#{key} must be symbol") { @dv.send("#{key}=".to_sym, "foo") }
      assert_nothing_raised { @dv.send("#{key}=".to_sym, value) }
    end
  end
  
  def test_formula1
    assert_raise(ArgumentError) { @dv.formula1 = 10 }
    assert_nothing_raised { @dv.formula1 = "=SUM(A1:A1)" }
    assert_equal(@dv.formula1, "=SUM(A1:A1)")
  end
  
  def test_formula2
    assert_raise(ArgumentError) { @dv.formula2 = 10 }
    assert_nothing_raised { @dv.formula2 = "=SUM(A1:A1)" }
    assert_equal(@dv.formula2, "=SUM(A1:A1)")
  end
  
  def test_allowBlank
    assert_raise(ArgumentError) { @dv.allowBlank = "foo´" }
    assert_nothing_raised { @dv.allowBlank = false }
    assert_equal(@dv.allowBlank, false)
  end
  
  def test_error
    assert_raise(ArgumentError) { @dv.error = :symbol }
    assert_nothing_raised { @dv.error = "This is a error message" }
    assert_equal(@dv.error, "This is a error message")
  end
  
  def test_errorStyle
    assert_raise(ArgumentError) { @dv.errorStyle = "foo" }
    assert_nothing_raised { @dv.errorStyle = :information }
    assert_equal(@dv.errorStyle, :information)
  end
  
  def test_errorTitle
    assert_raise(ArgumentError) { @dv.errorTitle = :symbol }
    assert_nothing_raised { @dv.errorTitle = "This is the error title" }
    assert_equal(@dv.errorTitle, "This is the error title")
  end
  
  def test_operator
    assert_raise(ArgumentError) { @dv.operator = "foo" }
    assert_nothing_raised { @dv.operator = :greaterThan }
    assert_equal(@dv.operator, :greaterThan)
  end
  
  def test_prompt
    assert_raise(ArgumentError) { @dv.prompt = :symbol }
    assert_nothing_raised { @dv.prompt = "This is a prompt message" }
    assert_equal(@dv.prompt, "This is a prompt message")
  end
  
  def test_promptTitle
    assert_raise(ArgumentError) { @dv.promptTitle = :symbol }
    assert_nothing_raised { @dv.promptTitle = "This is the prompt title" }
    assert_equal(@dv.promptTitle, "This is the prompt title")
  end
  
  def test_showDropDown
    assert_raise(ArgumentError) { @dv.showDropDown = "foo´" }
    assert_nothing_raised { @dv.showDropDown = false }
    assert_equal(@dv.showDropDown, false)
  end
  
  def test_showErrorMessage
    assert_raise(ArgumentError) { @dv.showErrorMessage = "foo´" }
    assert_nothing_raised { @dv.showErrorMessage = false }
    assert_equal(@dv.showErrorMessage, false)
  end
  
  def test_showInputMessage
    assert_raise(ArgumentError) { @dv.showInputMessage = "foo´" }
    assert_nothing_raised { @dv.showInputMessage = false }
    assert_equal(@dv.showInputMessage, false)
  end
  
  def test_sqref
    assert_raise(ArgumentError) { @dv.sqref = 10 }
    assert_nothing_raised { @dv.sqref = "A1:A1" }
    assert_equal(@dv.sqref, "A1:A1")
  end
  
  def test_type
    assert_raise(ArgumentError) { @dv.type = "foo" }
    assert_nothing_raised { @dv.type = :list }
    assert_equal(@dv.type, :list)
  end
  
  def test_whole_decimal_data_time_textLength_to_xml
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet :name=>"data_validation"
    @ws.add_data_validation("A1", { :type => :whole, :operator => :between, :formula1 => '5', :formula2 => '10', 
        :showErrorMessage => true, :errorTitle => 'Wrong input', :error => 'Only values between 5 and 10', 
        :errorStyle => :information, :showInputMessage => true, :promptTitle => 'Be carful!', 
        :prompt => 'Only values between 5 and 10'})
    
    doc = Nokogiri::XML.parse(@ws.to_xml_string)
    
    #test attributes
    assert_equal(1, doc.xpath("//xmlns:worksheet/xmlns:dataValidations[@count='1']/xmlns:dataValidation[@sqref='A1']
      [@promptTitle='Be carful!'][@prompt='Only values between 5 and 10'][@operator='between'][@errorTitle='Wrong input']
      [@error='Only values between 5 and 10'][@showErrorMessage=1][@allowBlank=1][@showInputMessage=1][@type='whole']
      [@errorStyle='information']").size)
    assert doc.xpath("//xmlns:worksheet/xmlns:dataValidations[@count='1']/xmlns:dataValidation[@sqref='A1']
      [@promptTitle='Be carful!'][@prompt='Only values between 5 and 10'][@operator='between'][@errorTitle='Wrong input']
      [@error='Only values between 5 and 10'][@showErrorMessage=1][@allowBlank=1][@showInputMessage=1]
      [@type='whole'][@errorStyle='information']")
    
    #test forumula1
    assert_equal(1, doc.xpath("//xmlns:worksheet/xmlns:dataValidations/xmlns:dataValidation/xmlns:formula1").size)
    assert doc.xpath("//xmlns:worksheet/xmlns:dataValidations/xmlns:dataValidation/xmlns:formula1='5'")
    
    #test forumula2
    assert_equal(1, doc.xpath("//xmlns:worksheet/xmlns:dataValidations/xmlns:dataValidation/xmlns:formula2").size)
    assert doc.xpath("//xmlns:worksheet/xmlns:dataValidations/xmlns:dataValidation/xmlns:formula2='10'")
  end
  
  def test_list_to_xml
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet :name=>"data_validation"
    @ws.add_data_validation("A1", { :type => :list, :formula1 => 'A1:A5',
        :showErrorMessage => true, :errorTitle => 'Wrong input', :error => 'Only values from list', 
        :errorStyle => :stop, :showInputMessage => true, :promptTitle => 'Be carful!', 
        :prompt => 'Only values from list', :showDropDown => true})
    
    doc = Nokogiri::XML.parse(@ws.to_xml_string)
    
    #test attributes
    assert_equal(1, doc.xpath("//xmlns:worksheet/xmlns:dataValidations[@count='1']/xmlns:dataValidation[@sqref='A1']
      [@promptTitle='Be carful!'][@prompt='Only values from list'][@errorTitle='Wrong input'][@error='Only values from list']
      [@showErrorMessage=1][@allowBlank=1][@showInputMessage=1][@showDropDown=1][@type='list']
      [@errorStyle='stop']").size)
    assert doc.xpath("//xmlns:worksheet/xmlns:dataValidations[@count='1']/xmlns:dataValidation[@sqref='A1']
      [@promptTitle='Be carful!'][@prompt='Only values from list'][@errorTitle='Wrong input'][@error='Only values from list']
      [@showErrorMessage=1][@allowBlank=1][@showInputMessage=1][@showDropDown=1][@type='list'][@errorStyle='stop']")
    
    #test forumula1
    assert_equal(1, doc.xpath("//xmlns:worksheet/xmlns:dataValidations/xmlns:dataValidation/xmlns:formula1").size)
    assert doc.xpath("//xmlns:worksheet/xmlns:dataValidations/xmlns:dataValidation/xmlns:formula1='A1:A5'")
  end
  
  def test_custom_to_xml
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet :name=>"data_validation"
    @ws.add_data_validation("A1", { :type => :custom, :formula1 => '=5/2',
        :showErrorMessage => true, :errorTitle => 'Wrong input', :error => 'Only values corresponding formula', 
        :errorStyle => :stop, :showInputMessage => true, :promptTitle => 'Be carful!', 
        :prompt => 'Only values corresponding formula'})
    
    doc = Nokogiri::XML.parse(@ws.to_xml_string)
    
    #test attributes
    assert_equal(1, doc.xpath("//xmlns:worksheet/xmlns:dataValidations[@count='1']/xmlns:dataValidation[@sqref='A1'][@promptTitle='Be carful!']
      [@prompt='Only values corresponding formula'][@errorTitle='Wrong input'][@error='Only values corresponding formula'][@showErrorMessage=1]
      [@allowBlank=1][@showInputMessage=1][@type='custom'][@errorStyle='stop']").size)
    assert doc.xpath("//xmlns:worksheet/xmlns:dataValidations[@count='1']/xmlns:dataValidation[@sqref='A1'][@promptTitle='Be carful!']
      [@prompt='Only values corresponding formula'][@errorTitle='Wrong input'][@error='Only values corresponding formula']
      [@showErrorMessage=1][@allowBlank=1][@showInputMessage=1][@type='custom'][@errorStyle='stop']")
    
    #test forumula1
    assert_equal(1, doc.xpath("//xmlns:worksheet/xmlns:dataValidations/xmlns:dataValidation/xmlns:formula1").size)
    assert doc.xpath("//xmlns:worksheet/xmlns:dataValidations/xmlns:dataValidation/xmlns:formula1='=5/2'")
  end
  
  def test_multiple_datavalidations_to_xml
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet :name=>"data_validation"
    @ws.add_data_validation("A1", { :type => :whole, :operator => :between, :formula1 => '5', :formula2 => '10', 
        :showErrorMessage => true, :errorTitle => 'Wrong input', :error => 'Only values between 5 and 10', 
        :errorStyle => :information, :showInputMessage => true, :promptTitle => 'Be carful!', 
        :prompt => 'Only values between 5 and 10'})
    @ws.add_data_validation("B1", { :type => :list, :formula1 => 'A1:A5',
        :showErrorMessage => true, :errorTitle => 'Wrong input', :error => 'Only values from list', 
        :errorStyle => :stop, :showInputMessage => true, :promptTitle => 'Be carful!', 
        :prompt => 'Only values from list', :showDropDown => true})
    
    doc = Nokogiri::XML.parse(@ws.to_xml_string)
    
    #test attributes
    assert_equal(1, doc.xpath("//xmlns:worksheet/xmlns:dataValidations[@count='2']/xmlns:dataValidation[@sqref='A1']
      [@promptTitle='Be carful!'][@prompt='Only values between 5 and 10'][@operator='between'][@errorTitle='Wrong input']
      [@error='Only values between 5 and 10'][@showErrorMessage=1][@allowBlank=1][@showInputMessage=1][@type='whole']
      [@errorStyle='information']").size)
    assert doc.xpath("//xmlns:worksheet/xmlns:dataValidations[@count='2']/xmlns:dataValidation[@sqref='A1']
      [@promptTitle='Be carful!'][@prompt='Only values between 5 and 10'][@operator='between'][@errorTitle='Wrong input']
      [@error='Only values between 5 and 10'][@showErrorMessage=1][@allowBlank=1][@showInputMessage=1]
      [@type='whole'][@errorStyle='information']")
    
    #test attributes
    assert_equal(1, doc.xpath("//xmlns:worksheet/xmlns:dataValidations[@count='2']/xmlns:dataValidation[@sqref='B1']
      [@promptTitle='Be carful!'][@prompt='Only values from list'][@errorTitle='Wrong input'][@error='Only values from list']
      [@showErrorMessage=1][@allowBlank=1][@showInputMessage=1][@showDropDown=1][@type='list']
      [@errorStyle='stop']").size)
    assert doc.xpath("//xmlns:worksheet/xmlns:dataValidations[@count='2']/xmlns:dataValidation[@sqref='B1']
      [@promptTitle='Be carful!'][@prompt='Only values from list'][@errorTitle='Wrong input'][@error='Only values from list']
      [@showErrorMessage=1][@allowBlank=1][@showInputMessage=1][@showDropDown=1][@type='list'][@errorStyle='stop']")
  end
  
  def test_empty_attributes
     v = Axlsx::DataValidation.new
     assert_equal(nil, v.send(:get_valid_attributes))

  end
end
