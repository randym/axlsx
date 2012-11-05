module Axlsx
  # This module defines a single method for parsing options in class
  # initializers.
  module OptionsParser

    # Parses an options hash by calling any defined method by the same
    # name of the key postfixed with an '='
    # @param [Hash] options Options to parse.
    def parse_options(options={})
      options.each do |key, value|
        self.send("#{key}=", value) if self.respond_to?("#{key}=") && value != nil
      end
    end
  end
end
