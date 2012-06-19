# encoding: UTF-8
require 'tc_helper.rb'

class TestContentType < Test::Unit::TestCase
  def setup
    @package = Axlsx::Package.new
    @doc = Nokogiri::XML(@package.send(:content_types).to_xml_string)
  end

  def test_valid_document
    schema = Nokogiri::XML::Schema(File.open(Axlsx::CONTENT_TYPES_XSD))
    assert(schema.validate(@doc).map{ |e| puts e.message; e.message }.empty?)
  end

  def test_pre_built_types

    o_path = "//xmlns:Override[@ContentType='%s']"
    d_path = "//xmlns:Default[@ContentType='%s']"

    #default
    assert_equal(@doc.xpath("//xmlns:Default").size, 2, "There should be 2 default types")

    node = @doc.xpath(d_path % Axlsx::XML_CT).first
    assert_equal(node["Extension"], "#{Axlsx::XML_EX}", "xml content type invalid")

    node = @doc.xpath(d_path % Axlsx::RELS_CT).first
    assert_equal(node["Extension"],"#{Axlsx::RELS_EX}", "relationships content type invalid")

    #overrride
    assert_equal(@doc.xpath("//xmlns:Override").size, 4, "There should be 4 Override types")

    node = @doc.xpath(o_path % Axlsx::APP_CT).first
    assert_equal(node["PartName"], "/#{Axlsx::APP_PN}", "App part name invalid")

    node = @doc.xpath(o_path % Axlsx::CORE_CT).first
    assert_equal(node["PartName"], "/#{Axlsx::CORE_PN}", "Core part name invalid")

    node = @doc.xpath(o_path % Axlsx::STYLES_CT).first
    assert_equal(node["PartName"], "/xl/#{Axlsx::STYLES_PN}", "Styles part name invalid")

    node = @doc.xpath(o_path % Axlsx::WORKBOOK_CT).first
    assert_equal(node["PartName"], "/#{Axlsx::WORKBOOK_PN}", "Workbook part invalid")
  end

  def test_should_get_worksheet_for_worksheets
    o_path = "//xmlns:Override[@ContentType='%s']"

    ws = @package.workbook.add_worksheet
    doc = Nokogiri::XML(@package.send(:content_types).to_xml_string)
    assert_equal(doc.xpath("//xmlns:Override").size, 5, "adding a worksheet should add another type")
    assert_equal(doc.xpath(o_path % Axlsx::WORKSHEET_CT).last["PartName"], "/xl/#{ws.pn}", "Worksheet part invalid")

    ws = @package.workbook.add_worksheet
    doc = Nokogiri::XML(@package.send(:content_types).to_xml_string)
    assert_equal(doc.xpath("//xmlns:Override").size, 6, "adding workship should add another type")
    assert_equal(doc.xpath(o_path % Axlsx::WORKSHEET_CT).last["PartName"], "/xl/#{ws.pn}", "Worksheet part invalid")

  end

  def test_drawings_and_charts_need_content_types
    o_path = "//xmlns:Override[@ContentType='%s']"
    ws = @package.workbook.add_worksheet

    c = ws.add_chart Axlsx::Pie3DChart
    doc = Nokogiri::XML(@package.send(:content_types).to_xml_string)
    assert_equal(doc.xpath("//xmlns:Override").size, 7, "expected 7 types got #{doc.css("Types Override").size}")
    assert_equal(doc.xpath(o_path % Axlsx::DRAWING_CT).first["PartName"], "/xl/#{ws.drawing.pn}", "Drawing part name invlid")
    assert_equal(doc.xpath(o_path % Axlsx::CHART_CT).last["PartName"], "/xl/#{c.pn}", "Chart part name invlid")

    c = ws.add_chart Axlsx::Pie3DChart
    doc = Nokogiri::XML(@package.send(:content_types).to_xml_string)
    assert_equal(doc.xpath("//xmlns:Override").size, 8, "expected 7 types got #{doc.css("Types Override").size}")
    assert_equal(doc.xpath(o_path % Axlsx::CHART_CT).last["PartName"], "/xl/#{c.pn}", "Chart part name invlid")
  end

end
