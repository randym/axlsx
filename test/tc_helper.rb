$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'simplecov'
SimpleCov.start do
  add_filter "/test/"
  add_filter "/vendor/"
end

require 'test/unit'
require "timecop"
require "axlsx.rb"

# Make sure all valid rIds are > 1000 - this should help catching the cases where rId is still
# determined by looking at the index of an object in an array etc.
Axlsx::Relationship.instance_variable_set :@next_free_id_counter, 1000 
