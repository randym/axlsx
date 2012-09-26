require 'tc_helper.rb'

class TestFilterColumn < Test::Unit::TestCase

  def setup
    @filter_column = Axlsx::FilterColumn.new(0, :filters, :filter_items => [200])
  end


  def test_initialize_col_id
    assert_raise ArgumentError do
      Axlsx::FilterColumn.new(0, :bobs_house_of_filter)
    end 
    assert_raise ArgumentError do 
      Axlsx::FilterColumn.new(:penut, :filters)
    end 
  end

  def test_initailize_filter_type
    assert @filter_column.filter.is_a?(Axlsx::Filters)
    assert_equal 1, @filter_column.filter.filter_items.size
  end

  def test_initialize_filter_type_filters_with_options
    assert_equal 200, @filter_column.filter.filter_items.first.val
  end

  def test_initialize_with_block
    filter_column = Axlsx::FilterColumn.new(0, :filters) do |filters|
      filters.filter_items = [700, 100, 5]
    end
    assert_equal 3, filter_column.filter.filter_items.size
    assert_equal 700, filter_column.filter.filter_items.first.val
    assert_equal 5, filter_column.filter.filter_items.last.val
  end

  def test_default_show_button
    assert_equal true, @filter_column.show_button
  end

  def test_default_hidden_button
    assert_equal false, @filter_column.hidden_button
  end

  def test_show_button
    assert_raise ArgumentError do
      @filter_column.show_button = :foo 
    end
    assert_nothing_raised { @filter_column.show_button = false }
  end

  def test_hidden_button
    assert_raise ArgumentError do
      @filter_column.hidden_button = :hoge
    end 
    assert_nothing_raised { @filter_column.hidden_button = true }
  end

  def test_col_id=
    assert_raise ArgumentError do 
    @filter_column.col_id = :bar
    end 
  assert_nothing_raised { @filter_column.col_id = 7 }
  end

  def test_to_xml_string
    doc = Nokogiri::XML(@filter_column.to_xml_string)
    assert doc.xpath("//filterColumn[@colId=#{@filter_column.col_id}]")
    assert doc.xpath("//filterColumn[@hiddenButton=#{@filter_column.hidden_button}]")
    assert doc.xpath("//filterColumn[@showButton=#{@filter_column.show_button}]")



    assert doc.xpath("//filterColumn/filters")
  end
end
