require 'axlsx/xml_dig_sig/transform.rb'
require 'axlsx/xml_dig_sig/reference'
require 'axlsx/xml_dig_sig/manifest'


module Axlsx
  class XmlDigSig

    # @TODO add Themes if we ever implement them.
    SIGNABLE_CONTENT_TYPES = [RELS_CT, STYLES_CT, WORKBOOK_CT, DRAWING_CT, TABLE_CT, COMMENT_CT, CHART_CT, SHARED_STRING_CT, WORKSHEET_CT]

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

    attr_reader :package, :certificate, :objects
    # serializes the signature
    def to_xml_string(str)

    end

    private
    def build_manifest
      # Not happy that this is private.
      # Makes me think we should be sending the sig object the package....
      manifest = Manifest.new
      parts = @package.send(:parts)
      parts.each do |part|
        if is_signable_part(part)
          manifiest.add_reference_for_package_part part
        end
      end
      
      @objects << SignatureObject.new(:id => 'idPackageObject', :content=> manifest)

      # can this be done?
      #@objects << Manifest.new do
      #  if is_signable_part(part)
      #    add_reference_for_package_part(part)
      #  end
      #end
    end
    # Generates a base64 encoded SHA1 hash for a canonicalized part entry
    def collect_references_from_package

      pacakge.send(:parts).each do |part|
        reference = Reference.new(part)
      end
    end

    private 
    def is_signable_part(part)
      SIGNABLE_CONETNET_TYPES.inclue? part.content_type
    end
  end
end
