require File.expand_path(File.dirname(__FILE__) + '/lib/axlsx/version.rb')
Gem::Specification.new do |s|
  s.name        = 'axlsx'
  s.rubyforge_project = 'axlsx'
  s.version     = Axlsx::VERSION
  s.date        = Time.now.strftime('%Y-%m-%d')
  s.summary     = "Author fully validated xlsx files with custom charts and styles"	
  s.platform    = Gem::Platform::RUBY       	     	  
  s.description = <<-eof
    axslx allows you to create Office Open XML Spreadsheet documents.
    It supports automated column widths, multiple worksheets, custom styles, cusomizable charts and allows you to validate your xlsx package before serialization. 
  eof
  s.author	= "Randy Morgan"
  s.email       = 'digital.ipeseity@gmail.com'
  s.homepage 	= 'https://github.com/randym/axlsx'
  s.has_rdoc    = 'axlsx'
  s.files 	= Dir.glob("{doc,lib,test,schema,examples}/**/*") + ['LICENSE','README.md','Rakefile','CHANGELOG.md']
  s.add_runtime_dependency 'nokogiri', '~> 1'
  s.add_runtime_dependency 'active_support', '~> 3'
  s.add_runtime_dependency 'rmagick', '~> 2.12'
  s.add_runtime_dependency 'rubyzip', '~> 0.9.4'
end
