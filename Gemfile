source "https://rubygems.org"
gemspec

group :test do
  gem "rake", "0.8.7"  if RUBY_VERSION == "1.9.2"
  gem "rake", ">= 0.8.7" unless RUBY_VERSION == "1.9.2"
end
group :profile do
  gem 'ruby-prof'
end
