require File.expand_path(File.dirname(__FILE__) + '/lib/axlsx/version.rb')

task :build => :gendoc do
  system "gem build axlsx.gemspec"
end

task :benchmark do
  require File.expand_path(File.dirname(__FILE__) + '/test/benchmark.rb')
end

task :gendoc do
  system "yardoc"
  system "yard stats --list-undoc"
end


require 'rake/testtask'
Rake::TestTask.new do |t|
  t.libs << 'test'
  t.test_files = FileList['test/**/tc_*.rb']
  t.verbose = false
  t.warning = true
end

task :release => :build do
  system "gem push caxlsx-#{Axlsx::VERSION}.gem"
end

task :default => :test
