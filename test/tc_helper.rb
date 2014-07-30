$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'simplecov'
SimpleCov.start do
  add_filter "/test/"
  add_filter "/vendor/"
end

require 'minitest/autorun'
require "timecop"
require "axlsx.rb"

# Mock assert_nothing_raised for now
# See http://blog.zenspider.com/blog/2012/01/assert_nothing_tested.html
def assert_nothing_raised(*args, &block)
  yield if block_given?
end
