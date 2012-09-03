module Axlsx
  class OfficeObject < SignatureObject
    def initialize(axlsx_package)
      super :id => 'idOfficeObject'
      @content = [signature_properties]
    end

    def signature_properties
      unless @signature_properties
        signature_info = SignatureInfo.new
        signature_info_property = SignatureProperty.new(:id => "idOfficeV1Details", :target => "idPackageSignature")
        signature_info_property.content << signature_info
        @signature_properties = SignatureProperties.new
        @signature_properties << signature_info_property
      end
      @signature_properties
    end

  end
end

