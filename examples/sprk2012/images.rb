require 'axlsx'
package = Axlsx::Package.new
workbook = package.workbook

workbook.add_worksheet(:name => "Image with Hyperlink") do |sheet|
  img = File.expand_path('../image1.jpeg', __FILE__)
  sheet.add_image(:image_src => img, :noSelect => true, :noMove => true, :hyperlink=>"http://axlsx.blogspot.com") do |image|
    image.width=720
    image.height=666
    image.hyperlink.tooltip = "Labeled Link"
    image.start_at 2, 2
  end
end

