$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'
xls = Axlsx::Package.new
wb = xls.workbook
wb.add_worksheet do |ws|
  img = File.expand_path('../image1.jpeg', __FILE__)
  ws.add_image(:image_src => img) do |image|
    image.start_at 2, 2
    image.end_at 5, 5
  end
end
wb.add_worksheet do |ws|
  img = File.expand_path('../image1.jpeg', __FILE__)
  ws.add_image(:image_src => img, :start_at => "B2") do |image|
    image.width = 70
    image.height = 50
  end
end
wb.add_worksheet do |ws|
  img = File.expand_path('../image1.jpeg', __FILE__)
  ws.add_image(:image_src => img, :start_at => [1,1]) do |image|
    image.end_at "E7"
  end
end



xls.serialize 'anchor.xlsx'
