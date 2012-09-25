require 'tc_helper.rb'

class TestFilters < Test::Unit::TestCase
  def setup
    @filters = Axlsx::Filters.new(:filter_items => [1, 'a'], :date_group_items =>[ { :date_time_grouping => :year, :year => 2012 } ] , :blank => true)
  end

  def test_initialize
    assert_equal Axlsx::Filters::CALENDAR_TYPES.first, @filters.calendar_type
  end

  def blank
    assert_equal false, @filters.blank
    assert_raise(ArgumentError) { @filters.blank = :only_if_you_want_it }
    @filters.blank = true
    assert_equal true, @filters.blank
  end

  def test_calendar_type
    assert_raise(ArgumentError) { @filters.calendar_type = 'monkey calendar' }
    @filters.calendar_type = 'japan'
    assert_equal('japan', @filters.calendar_type)
  end

  def test_filters_items
    assert @filters.filter_items.is_a?(Array)
    assert_equal 2, @filters.filter_items.size
  end

  def test_date_group_items
    assert @filters.date_group_items.is_a?(Array)
    assert_equal 1, @filters.date_group_items.size
  end

end

