require 'tc_helper.rb'

class TestPageSetup < Test::Unit::TestCase

  def setup
    @p = Axlsx::Package.new
    ws = @p.workbook.add_worksheet :name => "hmmm"
    @ps = ws.page_setup
  end

  def test_initialize
    assert_equal(nil, @ps.fit_to_height)
    assert_equal(nil, @ps.fit_to_width)
    assert_equal(nil, @ps.orientation)
    assert_equal(nil, @ps.paper_height)
    assert_equal(nil, @ps.paper_width)
    assert_equal(nil, @ps.scale)
  end

  def test_initialize_with_options
    page_setup = { :fit_to_height => 1,
                   :fit_to_width => 2,
                   :orientation => :landscape,
                   :paper_height => "297mm",
                   :paper_width => "210mm",
                   :scale => 50 }

    optioned = @p.workbook.add_worksheet(:name => 'optioned', :page_setup => page_setup).page_setup
    assert_equal(1, optioned.fit_to_height)
    assert_equal(2, optioned.fit_to_width)
    assert_equal(:landscape, optioned.orientation)
    assert_equal("297mm", optioned.paper_height)
    assert_equal("210mm", optioned.paper_width)
    assert_equal(50, optioned.scale)
  end

  def test_set_all_values
    @ps.set(:fit_to_height => 1, :fit_to_width => 2, :orientation => :landscape, :paper_height => "297mm", :paper_width => "210mm", :scale => 50)
    assert_equal(1, @ps.fit_to_height)
    assert_equal(2, @ps.fit_to_width)
    assert_equal(:landscape, @ps.orientation)
    assert_equal("297mm", @ps.paper_height)
    assert_equal("210mm", @ps.paper_width)
    assert_equal(50, @ps.scale)
  end

  def test_paper_size
    assert_raise(ArgumentError) { @ps.paper_size = 119 }
    assert_nothing_raised {  @ps.paper_size = 10 }
  end

  def test_set_some_values
    @ps.set(:fit_to_width => 2, :orientation => :portrait)
    assert_equal(2, @ps.fit_to_width)
    assert_equal(:portrait, @ps.orientation)
    assert_equal(nil, @ps.fit_to_height)
    assert_equal(nil, @ps.paper_height)
    assert_equal(nil, @ps.paper_width)
    assert_equal(nil, @ps.scale)
  end

  def test_default_fit_to_page?
    assert(@ps.fit_to_width == nil && @ps.fit_to_height == nil)
    assert(@ps.fit_to_page? == false)
  end

  def test_with_height_fit_to_page?
    assert(@ps.fit_to_width == nil && @ps.fit_to_height == nil)
    @ps.set(:fit_to_height => 1)
    assert(@ps.fit_to_page?)
  end

  def test_with_width_fit_to_page?
    assert(@ps.fit_to_width == nil && @ps.fit_to_height == nil)
    @ps.set(:fit_to_width => 1)
    assert(@ps.fit_to_page?)
  end

  def test_to_xml_all_values
    @ps.set(:fit_to_height => 1, :fit_to_width => 2, :orientation => :landscape, :paper_height => "297mm", :paper_width => "210mm", :scale => 50)
    doc = Nokogiri::XML.parse(@ps.to_xml_string)
    assert_equal(1, doc.xpath(".//pageSetup[@fitToHeight='1'][@fitToWidth='2'][@orientation='landscape'][@paperHeight='297mm'][@paperWidth='210mm'][@scale='50']").size)
  end

  def test_to_xml_some_values
    @ps.set(:orientation => :portrait)
    doc = Nokogiri::XML.parse(@ps.to_xml_string)
    assert_equal(1, doc.xpath(".//pageSetup[@orientation='portrait']").size)
    assert_equal(0, doc.xpath(".//pageSetup[@fitToHeight]").size)
    assert_equal(0, doc.xpath(".//pageSetup[@fitToWidth]").size)
    assert_equal(0, doc.xpath(".//pageSetup[@paperHeight]").size)
    assert_equal(0, doc.xpath(".//pageSetup[@paperWidth]").size)
    assert_equal(0, doc.xpath(".//pageSetup[@scale]").size)
  end

  def test_fit_to_height
    assert_raise(ArgumentError) { @ps.fit_to_height = 1.5 }
    assert_nothing_raised { @ps.fit_to_height = 2 }
    assert_equal(2, @ps.fit_to_height)
  end

  def test_fit_to_width
    assert_raise(ArgumentError) { @ps.fit_to_width = false }
    assert_nothing_raised { @ps.fit_to_width = 1 }
    assert_equal(1, @ps.fit_to_width)
  end

  def test_orientation
    assert_raise(ArgumentError) { @ps.orientation = "" }
    assert_nothing_raised { @ps.orientation = :default }
    assert_equal(:default, @ps.orientation)
  end

  def test_paper_height
    assert_raise(ArgumentError) { @ps.paper_height = 99 }
    assert_nothing_raised { @ps.paper_height = "11in" }
    assert_equal("11in", @ps.paper_height)
  end

  def test_paper_width
    assert_raise(ArgumentError) { @ps.paper_width = "22" }
    assert_nothing_raised { @ps.paper_width = "29.7cm" }
    assert_equal("29.7cm", @ps.paper_width)
  end

  def test_scale
    assert_raise(ArgumentError) { @ps.scale = 50.5 }
    assert_nothing_raised { @ps.scale = 99 }
    assert_equal(99, @ps.scale)
  end

  def test_fit_to
    fits = @ps.fit_to(:width => 1)
    assert_equal([1, 999], fits)
    fits = @ps.fit_to :height => 1
    assert_equal(fits, [999 ,1])
    fits = @ps.fit_to :height => 7, :width => 2
    assert_equal(fits, [2, 7])
    assert_raise(ArgumentError) { puts @ps.fit_to(:width => true)}


  end
end
