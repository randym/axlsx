module Axlsx
  class PackageObject < SignatureObject
    def initialize(axlsx_package)
      super :id => 'idPackageObject'
      @content = [axlsx_package.signature_manifest, signature_properties]
    end

    def signature_properties
      unless @signature_properties
        signature_time = SignatureTime.new(:time => Time.zone.now)
        signature_time_property = SignatureProperty.new(:id => "idSignatureTime", :target => "#idPackageSignature")
        signature_time_property.content << signature_time
        @signature_properties = SignatureProperties.new
        @signature_properties << signature_time_property
      end
      @signature_properties
    end

  end
end
