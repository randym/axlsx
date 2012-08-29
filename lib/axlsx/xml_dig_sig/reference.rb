module Axlsx
  class Reference
    require 'digest'
    require 'base64'

    # Digest method used in XML Digital Signatures
    @@digest_method = "http://www.w3.org/2000/09/xmldsig#sha1"

    # The reference class represents an XML Digital Signature reference object.
    # @param [String] part - The object to be referenced. This is usually parts of the package, or a reference in the signature it self.
    #                        I've no idea if you can sign binary files as well yet.
    #
    # @param [Hash] options options for type and transforms
    # @option [String] type an optional type property like http://uri.etsi.org/01903#SignedProperties which seems to only be applied to references of references.
    # @option [Array] tranforms a list transform classes to apply to the part. C14nTranform and RelationshipTransform are currently allowed
    def initialize(part, uri, options={})
      @part = part
      @uri = uri

    end

    # The document that will be transformed
    attr_reader :part

    # Tranforms to apply to this Reference's part before generating a digest value
    # @return [Array]
    attr_reader :transforms

    # The URI for this reference. For package parts this is the absolute path to the part followed by a 
    # query string defining the content type. 
    # e.g. /xl/styles.xml?ContentType=application/vnd.openxmlformats-officedocument.spreadsheetml.style
    # 
    # For signed info, it is a # prefixed id of one of the objects in the Signature.
    # e.g. #idPackageObject
    attr_accessor :uri

    # This is an optional attribute that I only see in SignedInfo references.
    # For signed Object nodes it is http://www.w3.org/2000/09/xmldsig#Object
    # For signed Properties nodes it is http://uri.etsi.org/01903#SignedProperties
    attr_accessor :type

    # Instantiates an array of transforms for this object. Not all references use tranforms.
    # Research indicates that non relation package parts only have the C14nTransform, and relation
    # parts have the RelationshipTransform applied, and then the C14Transform applied.
    # @param [Array] transform_classes An array containing C14nTransform and / or RelationshipTransform
    def transforms=(transform_classes)
      @tranforms = transform_classes.map do |kclass|
        RestrctionValidator.validate "Reference.transform", [C14nTransform, RelationshipTransform], kclass
        kclass.new
      end
    end

    def apply_tranforms
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

    def create_transform(name)
      class_name = Axlsx.camel("#{name}_transform")
      Axlsx.const_get(class_name)
    end

    def digest_value
      apply_transforms
      Base64.encode64(Digest::SHA1.digest(@doc))
    end
  end
end
