require 'rake'
require File.expand_path(File.dirname(__FILE__) + '/lib/axlsx/version.rb')
Gem::Specification.new do |s|
  s.name        = 'axlsx'
  s.version     = Axlsx::VERSION
  s.author	= "Randy Morgan"
  s.email       = 'digital.ipseity@gmail.com'
  s.homepage 	= 'https://github.com/randym/axlsx'
  s.platform    = Gem::Platform::RUBY       	     	  
  s.date        = Time.now.strftime('%Y-%m-%d')
  s.summary     = "OOXML (xlsx) with charts, styles, images and autowidth columns."
  s.has_rdoc    = 'axlsx'
  s.description = <<-eof
    xlsx generation with charts, images, automated column width, customizable styles and full schema validation. Axlsx excels at helping you generate beautiful Office Open XML Spreadsheet documents without having to understand the entire ECMA specification. Check out the README for some examples of how easy it is. Best of all, you can validate your xlsx file before serialization so you know for sure that anything generated is going to load on your client's machine.
  eof
  # s.files 	= Dir.glob("{doc,lib,test,schema,examples}/**/*") + %w{ LICENSE README.md Rakefile CHANGELOG.md }

  s.files = FileList.new('*', 'lib/**/*', 'doc/**/*', 'test/**/*', 'schema/**/*', 'examples/**/*') do |fl|
    fl.exclude("*.*~")
    fl.exclude(".*")
    fl.exclude("todo")
    fl.exclude("*.gem")
    fl.exclude("*.xlsx")
  end

  s.add_runtime_dependency 'nokogiri', '>= 1.4.1'
  s.add_runtime_dependency 'activesupport', '>= 2.3.9'
  s.add_runtime_dependency 'i18n', '>= 0.6.0'
  s.add_runtime_dependency 'rmagick', '>= 2.12.2'
  s.add_runtime_dependency 'zip', '>= 2.0.2'
  s.add_development_dependency 'rake'
  s.add_development_dependency 'bundler'

  s.required_ruby_version = '>= 1.8.7'
  s.require_path = 'lib'
end
