require 'tc_helper.rb'
class TestSimpleTypedList < Test::Unit::TestCase
  def setup
    @list = Axlsx::SimpleTypedList.new Integer
  end

  def teardown
  end

  def test_type_is_a_class_or_array_of_class
    assert_nothing_raised { Axlsx::SimpleTypedList.new Integer }
    assert_nothing_raised { Axlsx::SimpleTypedList.new [Integer,String] }
    assert_raise(ArgumentError) { Axlsx::SimpleTypedList.new }
    assert_raise(ArgumentError) { Axlsx::SimpleTypedList.new "1" }
    assert_raise(ArgumentError) { Axlsx::SimpleTypedList.new [Integer, "Class"] }
  end

  def test_indexed_based_assignment
    #should not allow nil assignment
    assert_raise(ArgumentError) { @list[0] = nil }
    assert_raise(ArgumentError) { @list[0] = "1" }
    assert_nothing_raised { @list[0] = 1 }
  end

  def test_concat_assignment
    assert_raise(ArgumentError) { @list << nil }
    assert_raise(ArgumentError) { @list << "1" }
    assert_nothing_raised { @list << 1 }
  end

  def test_concat_should_return_index
    assert( @list.size == 0 )
    assert( @list << 1 == 0 )
    assert( @list << 2 == 1 )
    @list.delete_at 0
    assert( @list << 3 == 1 )
    assert( @list.index(2) == 0 )
  end

  def test_push_should_return_index
    assert( @list.push(1) == 0 )
    assert( @list.push(2) == 1 )
    @list.delete_at 0
    assert( @list.push(3) == 1 )
    assert( @list.index(2) == 0 )
  end

  def test_locking
    @list.push 1
    @list.push 2
    @list.push 3
    @list.lock

    assert_raise(ArgumentError) { @list.delete 1  }
    assert_raise(ArgumentError) { @list.delete_at 1 }
    assert_raise(ArgumentError) { @list.delete_at 2 }
    @list.push 4
    assert_nothing_raised { @list.delete_at 3 }
    @list.unlock
    #ignore garbage
    assert_nothing_raised { @list.delete 0 }
    assert_nothing_raised { @list.delete 9 }
  end
  
  def test_delete
    @list.push 1
    assert(@list.size == 1)
    @list.delete 1
    assert(@list.empty?)
  end

  def test_equality
    @list.push 1
    @list.push 2
    assert_equal(@list.to_ary, [1,2])
  end
end
