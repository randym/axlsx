module Axlsx
  class Reference
    require 'digest'
    require 'base64'

    @@digest_method = "http://www.w3.org/2000/09/xmldsig#sha1"

    # The reference class represents an XML Digital Signature reference object.
    # @param [Hash] part - a hash of entry and doc where entry is the relative location
    def initialize(part)
      @transforms = []
      if part.entry.match('.rels')
        @transforms << RelationshipTransform.new
      end
      @transforms << C14nTransform.new
      @part = part
    end

    attr_reader :transforms

    def apply_tranforms
      @doc = @part.doc
      @transforms.each { |transform| transform.apply(self) } 
    end

    def to_xml_string(str = "")
      str << '<Reference URI="' << uri << '">'
      unless @transforms.empty?
        str << "<Transforms>"
        @transforms.each { |transform| transform.to_xml_str(str) }
        str << "</Transforms>"
      end
      str << '<DigestMethod Algorithm="' << @@digest_method << '"/>'
      str << '<DigestValue>' << digest_value << '</DigestValue>'
      str << '<Reference>'
    end
    private

    def digest_value
      apply_transforms
      Base64.encode64(Digest::SHA1.digest(@doc))
    end
  end
end
