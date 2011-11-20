require 'test/unit'
require 'axlsx.rb'

class TestRelationships < Test::Unit::TestCase
  def setup    
  end
  
  def teardown    
  end  

  def test_type
    assert_raise(ArgumentError) { Axlsx::Relationship.new 'type', 'target' }
    assert_nothing_raised { Axlsx::Relationship.new Axlsx::WORKSHEET_R, 'target' }
  end

end
