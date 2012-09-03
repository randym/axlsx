module Axlsx
  class QualifyingProperties


    # TODO OpenSSL::X509::Certificate
    def to_xml_string(str = '')
      str << '<xd:QualifyingProperties Target="#idPackageSignature" xmlns:xd="http://uri.etsi.org/01903/v1.3.2#">
      <xd:SignedProperties Id="idSignedProperties">
        <xd:SignedSignatureProperties>
          <xd:SigningTime>2012-08-23T16:32:07Z</xd:SigningTime>
          <xd:SigningCertificate>
            <xd:Cert>
              <xd:CertDigest>
                <DigestMethod Algorithm="http://www.w3.org/2000/09/xmldsig#sha1"/>
                <DigestValue>rp68UDMAR9RK4iN6DHx4xc5NvqQ=</DigestValue>
              </xd:CertDigest>
              <xd:IssuerSerial>
                <X509IssuerName>L=PDX, O=SweetSpot, E=tyler@sweetspotdiabetes.com, CN=Tyler Blitz</X509IssuerName>
                <X509SerialNumber>26100337707895980454701360632724310342</X509SerialNumber>
              </xd:IssuerSerial>
            </xd:Cert>
          </xd:SigningCertificate>
          <xd:SignaturePolicyIdentifier>
            <xd:SignaturePolicyImplied/>
          </xd:SignaturePolicyIdentifier>
        </xd:SignedSignatureProperties>
      </xd:SignedProperties>
      <xd:UnsignedProperties>
      </xd:UnsignedProperties>
    </xd:QualifyingProperties>'
    end
  end
end
