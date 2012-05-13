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

task :test do
     require 'rake/testtask'
     Rake::TestTask.new do |t|
       t.libs << 'test'
       t.test_files = FileList['test/**/tc_*.rb']
       t.verbose = true
     end
end

task :report do
  require 'cover_me'
  CoverMe.complete!
end

task :release => :build do
  system "gem push axlsx-#{Axlsx::VERSION}.gem"
end

task :default => :test
