require 'tc_helper.rb'

class TestPackage < Test::Unit::TestCase
  def setup
    @package = Axlsx::Package.new
    ws = @package.workbook.add_worksheet
    ws.add_row ['yes', 'we', 'can']
    ws.add_comment :author => 'bob', :text => 'penny!', :ref => 'A1'
    chart = ws.add_chart Axlsx::Pie3DChart
    chart.add_series :data=>[1,2,3], :labels=>["a", "b", "c"]
    @fname = 'axlsx_test_serialization.xlsx'
    img = File.expand_path('../../examples/image1.jpeg', __FILE__)
    ws.add_image(:image_src => img, :noSelect => true, :noMove => true, :hyperlink=>"http://axlsx.blogspot.com") do |image|
      image.width=720
      image.height=666
      image.hyperlink.tooltip = "Labeled Link"
      image.start_at 2, 2
    end
  end

  def test_use_autowidth
    @package.use_autowidth = false
    assert(@package.workbook.use_autowidth == false)
  end

  def test_core_accessor
    assert_equal(@package.core, @package.instance_values["core"])
    assert_raise(NoMethodError) {@package.core = nil }
  end

  def test_app_accessor
    assert_equal(@package.app, @package.instance_values["app"])
    assert_raise(NoMethodError) {@package.app = nil }
  end

  def test_use_shared_strings
    assert_equal(@package.use_shared_strings, nil)
    assert_raise(ArgumentError) {@package.use_shared_strings 9}
    assert_nothing_raised {@package.use_shared_strings = true}
    assert_equal(@package.use_shared_strings, @package.workbook.use_shared_strings)
  end

  def test_default_objects_are_created
    assert(@package.instance_values["app"].is_a?(Axlsx::App), 'App object not created')
    assert(@package.instance_values["core"].is_a?(Axlsx::Core), 'Core object not created')
    assert(@package.workbook.is_a?(Axlsx::Workbook), 'Workbook object not created')
    assert(Axlsx::Package.new.workbook.worksheets.size == 0, 'Workbook should not have sheets by default')
  end

  def test_serialization
    fname = 'axlsx_test_serialization.xlsx'
    assert_nothing_raised do
      begin
         z= @package.serialize(@fname)
         zf = Zip::ZipFile.open(@fname)
         @package.send(:parts).each{ |part| zf.get_entry(part[:entry]) }
         File.delete(@fname)
      rescue Errno::EACCES
         puts "WARNING:: test_serialization requires write access."
      end
     end
   end

  def test_validation
    assert_equal(@package.validate.size, 0, @package.validate)
    #how to test for failure? the internal validations on the models are so strict I cant break anthing.....
  end

  def test_parts
    p = @package.send(:parts)
    #all parts have an entry
    assert_equal(p.select{ |part| part[:entry] =~ /_rels\/\.rels/ }.size, 1, "rels missing")
    assert_equal(p.select{ |part| part[:entry] =~ /docProps\/core\.xml/ }.size, 1, "core missing")
    assert_equal(p.select{ |part| part[:entry] =~ /docProps\/app\.xml/ }.size, 1, "app missing")
    assert_equal(p.select{ |part| part[:entry] =~ /xl\/_rels\/workbook\.xml\.rels/ }.size, 1, "workbook rels missing")
    assert_equal(p.select{ |part| part[:entry] =~ /xl\/workbook\.xml/ }.size, 1, "workbook missing")
    assert_equal(p.select{ |part| part[:entry] =~ /\[Content_Types\]\.xml/ }.size, 1, "content types missing")
    assert_equal(p.select{ |part| part[:entry] =~ /xl\/styles\.xml/ }.size, 1, "styles missin")
    assert_equal(p.select{ |part| part[:entry] =~ /xl\/drawings\/_rels\/drawing\d\.xml\.rels/ }.size, @package.workbook.drawings.size, "one or more drawing rels missing")
    assert_equal(p.select{ |part| part[:entry] =~ /xl\/drawings\/drawing\d\.xml/ }.size, @package.workbook.drawings.size, "one or more drawings missing")
    assert_equal(p.select{ |part| part[:entry] =~ /xl\/charts\/chart\d\.xml/ }.size, @package.workbook.charts.size, "one or more charts missing")
    assert_equal(p.select{ |part| part[:entry] =~ /xl\/worksheets\/sheet\d\.xml/ }.size, @package.workbook.worksheets.size, "one or more sheet missing")
    assert_equal(p.select{ |part| part[:entry] =~ /xl\/worksheets\/_rels\/sheet\d\.xml\.rels/ }.size, @package.workbook.worksheets.size, "one or more sheet rels missing")
    assert_equal(p.select{ |part| part[:entry] =~ /xl\/comments\d\.xml/ }.size, @package.workbook.worksheets.size, "one or more sheet rels missing")


    #no mystery parts
    assert_equal(p.size, 15)

  end

  def test_shared_strings_requires_part
    @package.use_shared_strings = true
    p = @package.send(:parts)
    assert_equal(p.select{ |part| part[:entry] =~/xl\/sharedStrings.xml/}.size, 1, "shared strings table missing")
  end

  def test_workbook_is_a_workbook
    assert @package.workbook.is_a? Axlsx::Workbook
  end

  def test_base_content_types
    ct = @package.send(:base_content_types)
    assert(ct.select { |c| c.ContentType == Axlsx::RELS_CT }.size == 1, "rels content type missing")
    assert(ct.select { |c| c.ContentType == Axlsx::XML_CT }.size == 1, "xml content type missing")
    assert(ct.select { |c| c.ContentType == Axlsx::APP_CT }.size == 1, "app content type missing")
    assert(ct.select { |c| c.ContentType == Axlsx::CORE_CT }.size == 1, "core content type missing")
    assert(ct.select { |c| c.ContentType == Axlsx::STYLES_CT }.size == 1, "styles content type missing")
    assert(ct.select { |c| c.ContentType == Axlsx::WORKBOOK_CT }.size == 1, "workbook content type missing")
    assert(ct.size == 6)
  end

  def test_content_type_added_with_shared_strings
    @package.use_shared_strings = true
    ct = @package.send(:content_types)
    assert(ct.select { |ct| ct.ContentType == Axlsx::SHARED_STRINGS_CT }.size == 1)
  end

  def test_name_to_indices
    assert(Axlsx::name_to_indices('A1') == [0,0])
    assert(Axlsx::name_to_indices('A100') == [0,99], 'needs to axcept rows that contain 0')
  end
end
