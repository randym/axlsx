require 'test/unit'
require 'axlsx.rb'

class TestPackage < Test::Unit::TestCase
  def setup    
    @package = Axlsx::Package.new
    ws = @package.workbook.add_worksheet
    chart = ws.add_chart Axlsx::Pie3DChart
    chart.add_series :data=>[1,2,3], :labels=>["a", "b", "c"]
  end

  def test_default_objects_are_created
    assert(@package.instance_values["app"].is_a?(Axlsx::App), 'App object not created')
    assert(@package.instance_values["core"].is_a?(Axlsx::Core), 'Core object not created')
    assert(@package.workbook.is_a?(Axlsx::Workbook), 'Workbook object not created')
    assert(Axlsx::Package.new.workbook.worksheets.size == 0, 'Workbook should not have sheets by default')
  end

  def test_serialization
    fname = 'test_serialization.xlsx'
    assert_nothing_raised do
      if File.writable?(fname)
        z= @package.serialize(fname)              
        zf = Zip::ZipFile.open(fname)
        @package.send(:parts).each{ |part| zf.get_entry(part[:entry]) }
        File.delete(fname)        
      else
        puts "Skipping write to disk as write permission is not granted to this user"
      end
    end    
  end
  
  def test_validation
    assert_equal(@package.validate.size, 0, @package.validate)
    #how to test for failure? the internal validations on the models are so strict I cant break anthing.....
  end

  def test_parts
    p = @package.send(:parts)
    p.each do |part|
      #all parts must have :doc, :entry, :schema
      assert(part.keys.size == 3 && part.keys.reject{ |k| [:doc, :entry, :schema].include? k}.empty?) 
    end
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

    #no mystery parts
    assert_equal(p.size, 12)

  end
  
  def test_workbook_is_a_workbook
    assert @package.workbook.is_a? Axlsx::Workbook
  end
end
