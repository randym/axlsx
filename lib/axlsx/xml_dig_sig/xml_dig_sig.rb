module Axlsx
  class XmlDigSig

    # Canonical XML http://www.w3.org/TR/2001/REC-xml-c14n-20010314
    # XML Signature Processing http://www.w3.org/TR/xmldsig-core/#sec-Manifest
    # http://www.cs.auckland.ac.nz/~pgut001/pubs/xmlsec.txt
    def initialize(package, certificate, options)
      self.package = package
      self.certificate = certificate
    end

    attr_reader :package, :certificate

    # serializes the signature
    def to_xml_string(str)

    end

    private

    # Generates a base64 encoded SHA1 hash for a canonicalized part entry
    def collect_references_from_package

      pacakge.send(:parts).each do |part|
        reference = Reference.new(part)
      end
    end

  end
end
