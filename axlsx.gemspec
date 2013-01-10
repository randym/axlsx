require File.expand_path('../lib/axlsx/version', __FILE__)

Gem::Specification.new do |s|
  s.name        = 'axlsx'
  s.version     = Axlsx::VERSION
  s.author	= "Randy Morgan"
  s.email       = 'digital.ipseity@gmail.com'
  s.homepage 	= 'https://github.com/randym/axlsx'
  s.platform    = Gem::Platform::RUBY
  s.date        = Time.now.strftime('%Y-%m-%d')
  s.summary     = "excel OOXML (xlsx) with charts, styles, images and autowidth columns."
  s.has_rdoc    = 'axlsx'
  s.description = <<-eof
    xlsx spreadsheet generation with charts, images, automated column width, customizable styles and full schema validation. Axlsx helps you create beautiful Office Open XML Spreadsheet documents ( Excel, Google Spreadsheets, Numbers, LibreOffice) without having to understand the entire ECMA specification. Check out the README for some examples of how easy it is. Best of all, you can validate your xlsx file before serialization so you know for sure that anything generated is going to load on your client's machine.
  eof
  s.files = Dir.glob("{lib/**/*,examples/**/*.rb,examples/**/*.jpeg}") + %w{ LICENSE README.md Rakefile CHANGELOG.md .yardopts }
  s.test_files  = Dir.glob("{test/**/*}")

  s.add_runtime_dependency 'nokogiri', '>= 1.4.1'
  s.add_runtime_dependency 'rubyzip', '>= 0.9.5'
  s.add_runtime_dependency "htmlentities", "~> 4.3.1"

# This has been removed until JRuby can support the native extensions for redcarpet or yard removes the dependency
  s.add_development_dependency 'yard'
  s.add_development_dependency 'kramdown'
  s.add_development_dependency 'simplecov'
  s.required_ruby_version = '>= 1.8.7'
  s.require_path = 'lib'
end
