source 'https://rubygems.org'
gemspec

group :test do
  gem 'rake', '>= 0.8.7'
  gem 'simplecov'
  gem 'nokogiri', '>= 1.4.1'
end

group :profile do
  gem 'ruby-prof'
end

platforms :rbx do
  gem 'rubysl'
  gem 'rubysl-test-unit'
  gem 'racc'
  gem 'rubinius-coverage',  '~> 2.0'
end

platform :mri do
  if RUBY_VERSION > "1.9.2"
    gem 'escape_utils', '~> 1.0.1' 
  end
end

