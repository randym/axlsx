# -*- coding: utf-8 -*-
require 'tc_helper.rb'

class TestOverride < Test::Unit::TestCase

  def test_initialization_requires_Extension_and_ContentType
    err = "requires PartName and ContentType options"
    assert_raise(ArgumentError, err) { Axlsx::Override.new }
    assert_raise(ArgumentError, err) { Axlsx::Override.new :PartName=>"xml" }
    assert_raise(ArgumentError, err) { Axlsx::Override.new :ContentType=>"asdf" }
    assert_nothing_raised {Axlsx::Override.new :PartName=>"foo", :ContentType=>Axlsx::CHART_CT}
  end

  def test_content_type_restriction
    assert_raise(ArgumentError, "requires known content type") { Axlsx::Override.new :ContentType=>"asdf" }
  end

  def test_to_xml
    type = Axlsx::Override.new :PartName=>"somechart.xml", :ContentType=>Axlsx::CHART_CT
    doc = Nokogiri::XML(type.to_xml_string)
    assert_equal(doc.xpath("Override[@ContentType='#{Axlsx::CHART_CT}']").size, 1)
    assert_equal(doc.xpath("Override[@PartName='somechart.xml']").size, 1)

  end


end
