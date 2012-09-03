require 'axlsx/xml_dig_sig/transform.rb'
require 'axlsx/xml_dig_sig/reference'
require 'axlsx/xml_dig_sig/manifest'


module Axlsx
  class XmlDigSig

    # References - 
    # Canonical XML http://www.w3.org/TR/2001/REC-xml-c14n-20010314
    # XML Signature Processing http://www.w3.org/TR/xmldsig-core/#sec-Manifest
    # Partial implementation (Java) http://www.cs.auckland.ac.nz/~pgut001/pubs/xmlsec.txt
    def initialize(package, certificate, options)
      @objects = []
      @package = package
      @certificate = certificate
      # Generate Objects
      # create idPackageObject
      # get package parts and create a manifest
      # create the signature properties node
      # create the idOfficeObject
      # create QualifyingProperties
      # Generate KeyInfo from certificate
      # Generate Signed Info from objects
      # Generate SignatureValue from SingedInfo and KeyInfo
    end

    attr_reader :package, :certificate

    def objects
      @object ||= [package_object, office_object, qualifying_properites]
    end
    # serializes the signature
    def to_xml_string(str)

    end

    private

    def package_object
      @package_object ||= PackageObject.new
    end
  end
end
