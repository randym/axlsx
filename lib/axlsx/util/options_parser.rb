module Axlsx
  module OptionsParser
    def parse_options(options={})
      options.each do |key, value|
        self.send("#{key}=", value) if self.respond_to?("#{key}=") && value != nil
      end
    end
  end
end
