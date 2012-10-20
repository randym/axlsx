require 'tc_helper.rb'

class TestCol < Test::Unit::TestCase

  def setup
    @col = Axlsx::Col.new 1, 1
  end

  def test_initialize
    options = { :width => 12, :collapsed => true, :hidden => true, :outline_level => 1, :phonetic => true, :style => 1} 

    col = Axlsx::Col.new 0, 0, options
    options.each{ |key, value| assert_equal(col.send(key.to_sym), value) }
  end

  def test_min_max_required
    assert_raise(ArgumentError, 'min and max must be specified when creating a new column') { Axlsx::Col.new }
    assert_raise(ArgumentError, 'min and max must be specified when creating a new column') { Axlsx::Col.new nil, nil }
    assert_nothing_raised { Axlsx::Col.new 1, 1 }
  end

  def test_bestFit
    assert_equal(@col.bestFit, nil)
    assert_raise(NoMethodError, 'bestFit is read only') { @col.bestFit = 'bob' }
    @col.width = 1.999
    assert_equal(@col.bestFit, true, 'bestFit should be true when width has been set')
  end

  def test_collapsed
    assert_equal(@col.collapsed, nil)
    assert_raise(ArgumentError, 'collapsed must be boolean(ish)') { @col.collapsed = 'bob' }
    assert_nothing_raised('collapsed must be boolean(ish)') { @col.collapsed = true }
  end

  def test_customWidth
    assert_equal(@col.customWidth, nil)
    @col.width = 3
    assert_raise(NoMethodError, 'customWidth is read only') { @col.customWidth = 3 }
    assert_equal(@col.customWidth, true, 'customWidth is true when width is set')
  end

  def test_hidden
    assert_equal(@col.hidden, nil)
    assert_raise(ArgumentError, 'hidden must be boolean(ish)') { @col.hidden = 'bob' }
    assert_nothing_raised(ArgumentError, 'hidden must be boolean(ish)') { @col.hidden = true }
  end

  def test_outlineLevel
    assert_equal(@col.outlineLevel, nil)
    assert_raise(ArgumentError, 'outline level cannot be negative') { @col.outlineLevel = -1 }
    assert_raise(ArgumentError, 'outline level cannot be greater than 7') { @col.outlineLevel = 8 }
    assert_nothing_raised('can set outlineLevel') { @col.outlineLevel = 1 }
  end

  def test_phonetic
    assert_equal(@col.phonetic, nil)
    assert_raise(ArgumentError, 'phonetic must be boolean(ish)') { @col.phonetic = 'bob' }
    assert_nothing_raised(ArgumentError, 'phonetic must be boolean(ish)') { @col.phonetic = true }
  end

  def test_to_xml_string
    @col.width = 100
    doc = Nokogiri::XML(@col.to_xml_string)
    assert_equal(1, doc.xpath("//col [@bestFit='#{@col.best_fit}']").size)
    assert_equal(1, doc.xpath("//col [@max=#{@col.max}]").size)
    assert_equal(1, doc.xpath("//col [@min=#{@col.min}]").size)
    assert_equal(1, doc.xpath("//col [@width=#{@col.width}]").size)
    assert_equal(1, doc.xpath("//col [@customWidth='#{@col.custom_width}']").size)
  end

  def test_style
    assert_equal(@col.style, nil)
    @col.style = 1
    assert_equal(@col.style, 1)
    #TODO check that the style specified is actually in the styles xfs collection
  end

end
