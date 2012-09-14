$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../../lib"

require 'axlsx'
package = Axlsx::Package.new do |package|
  package.workbook.add_worksheet(:name => "imagex") do |sheet|
    img_path = File.expand_path('../../image1.jpeg', __FILE__)
    sheet.add_image(:image_src => img_path, :width => 720, :height => 666, :start_at => [2,2])
  end
end.serialize 'images.xlsx'
