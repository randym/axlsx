# require 'tc_helper.rb'

# class TestCatAxisData < Test::Unit::TestCase

#   def setup
#     p = Axlsx::Package.new
#     @ws = p.workbook.add_worksheet
#     @chart = @ws.drawing.add_chart Axlsx::Bar3DChart
#     @series = @chart.add_series :labels=>["zero", "one", "two"]
#   end

#   def test_initialize
#     assert(@series.labels.is_a?Axlsx::SimpleTypedList)
#     assert_equal(@series.labels, ["zero", "one", "two"])
#   end

#   def test_to_xml_string
#     doc = Nokogiri::XML(@chart.to_xml_string)
#     assert_equal(doc.xpath("//c:cat/c:strRef/c:f").size,1)
#     assert_equal(doc.xpath("//c:strCache/c:ptCount[@val='#{@series.labels.size}']").size,1)
#     @series.labels.each_with_index do |label, index|
#       assert_equal(doc.xpath("//c:strCache/c:pt[@idx='#{index}']").size,1)
#       assert_equal(doc.xpath("//c:strCache/c:pt/c:v[text()='#{label}']").size,1)
#     end
#   end

# end
