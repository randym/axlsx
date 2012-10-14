module Axlsx
  module OptionsParser
    def parse_options(options={})
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
    end
  end
end
