require 'tc_helper.rb'

class TestBarSeries < Test::Unit::TestCase

  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet :name=>"hmmm"
    @chart = @ws.add_chart Axlsx::Bar3DChart, :title => "fishery"
    @series = @chart.add_series :data=>[0,1,2], :labels=>["zero", "one", "two"], :title=>"bob", :colors => ['FF0000', '00FF00', '0000FF'], :shape => :cone
  end

  def test_initialize
    assert_equal(@series.title.text, "bob", "series title has been applied")
    assert_equal(@series.data.class, Axlsx::NumDataSource, "data option applied")
    assert_equal(@series.shape, :cone, "series shape has been applied")
    assert(@series.data.is_a?(Axlsx::NumDataSource))
    assert(@series.labels.is_a?(Axlsx::AxDataSource))
  end

  def test_colors
    assert_equal(@series.colors.size, 3)
  end

  def test_shape
    assert_raise(ArgumentError, "require valid shape") { @series.shape = :teardropt }
    assert_nothing_raised("allow valid shape") { @series.shape = :box }
    assert(@series.shape == :box)
  end

  def test_to_xml_string
    doc = Nokogiri::XML(@chart.to_xml_string)
    @series.colors.each_with_index do |color, index|
      assert_equal(doc.xpath("//c:dPt/c:idx[@val='#{index}']").size,1)
      assert_equal(doc.xpath("//c:dPt/c:spPr/a:solidFill/a:srgbClr[@val='#{@series.colors[index]}']").size,1)
    end
  end
end
