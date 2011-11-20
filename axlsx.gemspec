require File.expand_path(File.dirname(__FILE__) + '/lib/axlsx/util/constants.rb')
Gem::Specification.new do |s|
  s.name        = 'axlsx'
  s.version     = Axlsx::VERSION
  s.date          = Time.now.strftime('%Y-%m-%d')
  s.summary     = "Author fully validated xlsx files with custom charts and styles"	
  s.platform      = Gem::Platform::RUBY       	     	  
  s.description = <<-EOF
    axslx allows you to easily and skillfully create Office Open XML Spreadsheet documents. It supports multiple worksheets, custom styles, cusomizable charts and allows you to validate your xlsx package before serialization. 
EOF
  s.authors     = ["Randy Morgan"]
  s.email       = 'digital.ipeseity@gmail.com'
  s.files = Dir.glob("{docs,lib,test,schema}/**/*") + ['LICENSE', 'README.md', 'Rakefile']
  s.has_rdoc = 'yard'
  s.homepage    = ""  
  s.add_runtime_dependency 'nokogiri', '~> 1'
  s.add_runtime_dependency 'active_support', '~> 3'
  s.add_runtime_dependency 'rmagick', '~> 2.12'
  s.add_runtime_dependency 'rubyzip', '~> 0.9.4'
end
