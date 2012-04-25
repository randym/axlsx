# encoding: UTF-8
module Axlsx

  # The version of the gem
  # When using bunle exec rake and referencing the gem on github or locally
  # it will use the gemspec, which preloads this constant for the gem's version.
  # We check to make sure that it has not already been loaded
  VERSION="1.1.2" unless defined? Axlsx::VERSION

end
