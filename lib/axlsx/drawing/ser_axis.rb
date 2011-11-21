module Axlsx
  #A SarAxis object defines a series axis
  class SerAxis < Axis

    # @return [Boolean]
    attr_accessor :tickLblSkip

    # @return [Boolean]
    attr_accessor :tickMarkSkip

    # Creates a new SerAxis object
    # @param [Integer] axId the id of this axis
    # @param [Integer] crossAx the id of the perpendicular axis
    # @option options [Symbol] axPos
    # @option options [Symbol] tickLblPos
    # @option options [Symbol] crosses
    # @option options [Boolean] tickLblSkip
    # @option options [Symbol] tickMarkSkip
    def initialize(axId, crossAx, options={})
      super(axId, crossAx, options)
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
    end 

    def tickLblSkip=(v) Axlsx::validate_boolean(v); @tickLblSkip = v; end
    def tickMarkSkip=(v) Axlsx::validate_boolean(v); @tickMarkSkip = v; end

    # Serializes the series axis
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml.send('c:serAx') {
        super(xml)
        xml.send('c:tickLblSkip', :val=>@tickLblSkip) unless @tickLblSkip.nil?
        xml.send('c:tickMarkSkip', :val=>@tickMarkSkip) unless @tickMarkSkip.nil?
      }
    end
  end
  

end
