module Axlsx

  class SignatureProperty

    def initialize(options={})
      @content = []
      options.each do |name, value|
        self.send("#{name}=", value) if self.respond_to? "#{name}="
      end
    end

    attr_accessor :id
    attr_accessor :target

    attr_accessor :content

    def to_xml_string(str = '')
      str << '<SignatureProperty Id="' << id << '" Target="' << target << '">'
      @content.each { |item| item.to_xml_string(str)  }
      str << '</SignaturePropery>'
    end
  end
end
