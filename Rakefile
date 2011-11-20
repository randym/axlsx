require 'rake/testtask'

# gems not loading? try this
#RUBYOPT="rubygems"
#export RUBYOPT

Rake::TestTask.new do |t|
  t.libs << 'test'
  t.test_files = FileList['test/**/tc_*.rb']
  t.verbose = true
end

task :default => :test