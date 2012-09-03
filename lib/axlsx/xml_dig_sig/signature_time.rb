module Axlsx

  class SignatureTime

    TIME_FORMAT = "YYYY-MM-DDThh:mm:ssTZD"
    TIME_FORMAT_RUBY = "%Y-%M-%dT%H:%m:%sZ"
 
    def initialize(time)
      @time = time
    end

    def to_xml_string(str = '')
      str << '<mdssi:SignatureTime>'
      str << "<mdssi:Format>#{TIME_FORMAT}</mdssi:Format>"
      str << "<mdssi:Value>#{@time.strftime(TIME_FORMAT_RUBY)}</mdssi:Value>"
      str << '</mdssi:SignatureTime>'
    end
  end
end
