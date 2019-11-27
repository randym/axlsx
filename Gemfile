# frozen_string_literal: true
source 'https://rubygems.org'
gemspec

group :test do
  if RUBY_VERSION.to_i < 2
    gem 'rake', '>= 0.8.7', '< 11'
    gem 'json', '< 2'
  else
    gem 'rake'
  end
  gem 'simplecov', '>= 0.14.1'
  gem 'test-unit'
end

group :profile do
  gem 'ruby-prof', :platforms => :ruby
end

platforms :rbx do
  gem 'rubysl'
  gem 'rubysl-test-unit'
  gem 'racc'
  gem 'rubinius-coverage', '~> 2.0'
end
