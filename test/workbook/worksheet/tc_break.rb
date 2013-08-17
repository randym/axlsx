require 'tc_helper.rb'

class TestBreak < Test::Unit::TestCase

  def setup
    p = Axlsx::Package.new
    ws = p.workbook.add_worksheet
    @break = Axlsx::Break.new(ws, {id: 1, min: 1, max: 10, man: true, pt: false})
  end

  def test_id
    assert_equal(1, @break.id)
    assert_raises ArgumentError do
      Axlsx::Break.new(:hoge, {id: -1})
    end
  end

  def test_min
  end

  def test_max
  end

  def test_man
  end

  def test_pt
  end

  def test_to_xml_string
  end
end
