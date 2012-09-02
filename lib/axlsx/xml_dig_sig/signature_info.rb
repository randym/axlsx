module Axlsx
  # http://msdn.microsoft.com/en-us/library/dd950680(v=office.12)
  # Absolutely nothing in the specs for this, so the MS reference is authoritative?
  class SignatureInfo

    def initialize(options={})

    end

    #TODO implement the elements that actually count.
    # thinking provider id, @rpvoder detail, signature type?
    def to_xml_string(str = '')
      str << '<SignatureInfoV1 xmlns="http://schemas.microsoft.com/office/2006/digsig">
          <SetupID/>
          <SignatureText/>
          <SignatureImage/>
          <SignatureComments>security</SignatureComments>
          <WindowsVersion>6.1</WindowsVersion>
          <OfficeVersion>14.0</OfficeVersion>
          <ApplicationVersion>14.0</ApplicationVersion>
          <Monitors>2</Monitors>
          <HorizontalResolution>1920</HorizontalResolution>
          <VerticalResolution>1058</VerticalResolution>
          <ColorDepth>32</ColorDepth>
          <SignatureProviderId>{00000000-0000-0000-0000-000000000000}</SignatureProviderId>
          <SignatureProviderUrl/>
          <SignatureProviderDetails>9</SignatureProviderDetails>
          <ManifestHashAlgorithm>http://www.w3.org/2000/09/xmldsig#sha1</ManifestHashAlgorithm>
          <SignatureType>1</SignatureType>
        </SignatureInfoV1>'
    end
  end
end
