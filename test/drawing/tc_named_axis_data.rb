# require 'tc_helper.rb'

# class TestNamedAxisData < Test::Unit::TestCase

#   def setup
#     p = Axlsx::Package.new
#     @ws = p.workbook.add_worksheet
#     @chart = @ws.drawing.add_chart Axlsx::Line3DChart
#     @series = @chart.add_series :data=>[0,1,2]
#   end

#   def test_initialize
#     assert(@series.data.is_a?Axlsx::SimpleTypedList)
#     assert_equal(@series.data, [0,1,2])
#   end

#   def test_to_xml_string
#     doc = Nokogiri::XML(@chart.to_xml_string)
#     assert_equal(doc.xpath("//c:val/c:numRef/c:f").size,1)
#     assert_equal(doc.xpath("//c:numCache/c:ptCount[@val='#{@series.data.size}']").size,1)
#     @series.data.each_with_index do |datum, index|
#       assert_equal(doc.xpath("//c:numCache/c:pt[@idx='#{index}']").size,1)
#       assert_equal(doc.xpath("//c:numCache/c:pt/c:v[text()='#{datum}']").size,1)
#     end
#   end

# end
