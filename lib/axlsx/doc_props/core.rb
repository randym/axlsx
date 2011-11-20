module Axlsx
  # The core object for the package.
  # @note Packages manage their own core object.
  # @see Package#core
  class Core
    # The author of the document. By default this is 'axlsx'
    # @return [String]
    attr_accessor :creator
    
    # Creates a new Core object.
    # @option options [String] creator
    def initialize(options={})
      @creator = options[:creator] || 'axlsx'      
    end

    # Serializes the core object. The created dcterms item is set to the current time when this method is called.
    # @return [String]
    def to_xml()
      builder = Nokogiri::XML::Builder.new(:encoding => ENCODING) do |xml|
        xml.send('cp:coreProperties', 
                 :"xmlns:cp" => CORE_NS, 
                 :'xmlns:dc' => CORE_NS_DC, 
                 :'xmlns:dcmitype'=>CORE_NS_DCMIT, 
                 :'xmlns:dcterms'=>CORE_NS_DCT, 
                 :'xmlns:xsi'=>CORE_NS_XSI) {
          xml['dc'].creator self.creator
          xml['dcterms'].created Time.now.strftime('%Y-%m-%dT%H:%M:%S'), :'xsi:type'=>"dcterms:W3CDTF"
          xml['cp'].revision 0
        }
      end  
      builder.to_xml
    end
  end
end
