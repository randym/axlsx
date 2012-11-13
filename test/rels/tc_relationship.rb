require 'tc_helper.rb'

class TestRelationships < Test::Unit::TestCase
  def setup
  end

  def teardown
  end

  def test_type
    assert_raise(ArgumentError) { Axlsx::Relationship.new 'type', 'target' }
    assert_nothing_raised { Axlsx::Relationship.new Axlsx::WORKSHEET_R, 'target' }
    assert_nothing_raised { Axlsx::Relationship.new Axlsx::COMMENT_R, 'target' }
  end

  def test_target_mode
    assert_raise(ArgumentError) { Axlsx::Relationship.new 'type', 'target', :target_mode => "FISH" }
    assert_nothing_raised { Axlsx::Relationship.new( Axlsx::WORKSHEET_R, 'target', :target_mode => :External) }
  end

  def test_ampersand_escaping_in_target
    r = Axlsx::Relationship.new(Axlsx::HYPERLINK_R, "http://example.com?foo=1&bar=2", :target_mod => :External)
    doc = Nokogiri::XML(r.to_xml_string(1))
    assert_equal(doc.xpath("//Relationship[@Target='http://example.com?foo=1&bar=2']").size, 1)
  end
end
