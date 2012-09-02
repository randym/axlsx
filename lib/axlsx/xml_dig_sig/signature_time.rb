module Axlsx

  class SignatureTime

    def initialize(time)
      @time = time
    end

    def to_xml_string(str = '')
      str << '<mdssi:SignatureTime>'
      str << '<mdssi:Format>YYYY-MM-DDThh:mm:ssTZD</mdssi:Format>'
      str << "<mdssi:Value>#{@time.strftime('%Y-%M-%d:%H:%m:%sZ')}</mdssi:Value>"
      str << '</mdssi:SignatureTime>'
    end
  end
end
