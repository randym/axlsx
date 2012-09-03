module Axlsx
  # The object class represents one of the signed objects in a 
  # DigitalSignature object.
  # @note This is serialized as "Object". The use of SignatureObject is only to 
  # distinguish between the existing ruby Object class.
  class SignatureObject
    # Creates a new ManifestObject
    # @param [Hash] options The options for this manifest object.
    # @option [String] id The Id of the object. At this point this *shoud* 
    #                     be one of idPackageObject or idOfficeObject
    # @option [Hash] namespace The namespace to apply to the object. 
    #                          The Manifest idPackageObject requires this to be
    #                          { :'xmlns:mdssi' => 'http://schemas.openxmlformats.org/package/2006/digital-signature' }
    #                          For now, if you specify that id, it will automatically be set.
    # @option [Any] content The content of the object. This should be some Axlsx object that needs to be included in the 
    #                       signature objects like manifes, signature_properties, or qualifying_properties.
    #                       This object must expose a to_xml_string method for serialization.
    def intialize(options={})
      @content = []
      if options[:id] && options[:id] == 'idPackageObject'
        @namespace = 'http://schemas.openxmlformats.org/package/2006/digital-signature'
      end
      options.each do |name, value|
        self.send("#{name}=", value) if self.respond_to? "#{name}="
      end
    end

    # The namespace for the object. This is only defined, and force fed
    #  (which means you cannot easily change it) for objects that use 
    #  an id of  "idPackageObject"
    # @return [String]
    attr_reader :namespace

    # The contest of the object. e.g. a manifest
    # @return [Any]
    attr_reader :content

    # The optional id for the object.
    # @return [String]
    attr_accessor :id

    def to_xml_string(str = '')
      str << "<Object"
      str << " Id=\"#{id}\"" if id
      str << " xmlns:mdssi=\"#{namespace}\"" if namespace 
      str << ">"
      content.each { |item| item.to_xml_string(str) }
      str << '</Object>'
    end
  end
end
