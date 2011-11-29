require File.expand_path(File.dirname(__FILE__) + '/lib/axlsx/version.rb')

task :build => :gendoc do
  system "gem build axlsx.gemspec"
end

task :gendoc do   
  system "yardoc"
end

task :test do 
     require 'rake/testtask'
     puts "The current time is: #{Time.now}"
 
     puts "1900 was: #{Time.local(1900,1,1)}"
     puts "1904 was: #{Time.local(1904,1,1)}"

     Rake::TestTask.new do |t|
       t.libs << 'test'
       t.test_files = FileList['test/**/tc_*.rb']
       t.verbose = true
     end
end

task :release => :build do
  system "gem push axlsx-#{Axlsx::VERSION}.gem"
end

task :default => :test