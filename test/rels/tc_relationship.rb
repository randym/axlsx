require 'tc_helper.rb'

class TestRelationships < Test::Unit::TestCase
  
  def test_instances_with_different_attributes_have_unique_ids
    rel_1 = Axlsx::Relationship.new(Object.new, Axlsx::WORKSHEET_R, 'target')
    rel_2 = Axlsx::Relationship.new(Object.new, Axlsx::COMMENT_R, 'foobar')
    assert_not_equal rel_1.Id, rel_2.Id
  end
  
  def test_instances_with_same_attributes_share_id
    source_obj = Object.new
    instance = Axlsx::Relationship.new(source_obj, Axlsx::WORKSHEET_R, 'target')
    assert_equal instance.Id, Axlsx::Relationship.new(source_obj, Axlsx::WORKSHEET_R, 'target').Id
  end
  
  def test_target_is_only_considered_for_same_attributes_check_if_target_mode_is_external
    source_obj = Object.new
    rel_1 = Axlsx::Relationship.new(source_obj, Axlsx::WORKSHEET_R, 'target')
    rel_2 = Axlsx::Relationship.new(source_obj, Axlsx::WORKSHEET_R, '../target')
    assert_equal rel_1.Id, rel_2.Id
    
    rel_3 = Axlsx::Relationship.new(source_obj, Axlsx::HYPERLINK_R, 'target', :target_mode => :External)
    rel_4 = Axlsx::Relationship.new(source_obj, Axlsx::HYPERLINK_R, '../target', :target_mode => :External)
    assert_not_equal rel_3.Id, rel_4.Id
  end
  
  def test_type
    assert_raise(ArgumentError) { Axlsx::Relationship.new nil, 'type', 'target' }
    assert_nothing_raised { Axlsx::Relationship.new nil, Axlsx::WORKSHEET_R, 'target' }
    assert_nothing_raised { Axlsx::Relationship.new nil, Axlsx::COMMENT_R, 'target' }
  end

  def test_target_mode
    assert_raise(ArgumentError) { Axlsx::Relationship.new nil, 'type', 'target', :target_mode => "FISH" }
    assert_nothing_raised { Axlsx::Relationship.new( nil, Axlsx::WORKSHEET_R, 'target', :target_mode => :External) }
  end

  def test_ampersand_escaping_in_target
    r = Axlsx::Relationship.new(nil, Axlsx::HYPERLINK_R, "http://example.com?foo=1&bar=2", :target_mod => :External)
    doc = Nokogiri::XML(r.to_xml_string)
    assert_equal(doc.xpath("//Relationship[@Target='http://example.com?foo=1&bar=2']").size, 1)
  end
end
