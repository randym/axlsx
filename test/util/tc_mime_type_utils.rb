require 'tc_helper.rb'
class TestMimeTypeUtils < Test::Unit::TestCase
  def setup
    @test_img = File.dirname(__FILE__) + "/../../examples/image1.jpeg"
  end

  def teardown
  end

  def test_mime_type_utils
    assert_equal(Axlsx::MimeTypeUtils::get_mime_type(@test_img), 'image/jpeg')
  end
end
