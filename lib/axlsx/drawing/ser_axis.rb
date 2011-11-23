module Axlsx
  #A SerAxis object defines a series axis
  class SerAxis < Axis

    # The number of tick lables to skip between labels
    # @return [Integer]
    attr_accessor :tickLblSkip

    # The number of tickmarks to be skipped before the next one is rendered.
    # @return [Boolean]
    attr_accessor :tickMarkSkip

    # Creates a new SerAxis object
    # @param [Integer] axId the id of this axis. Inherited
    # @param [Integer] crossAx the id of the perpendicular axis. Inherited
    # @option options [Symbol] axPos. Inherited
    # @option options [Symbol] tickLblPos. Inherited
    # @option options [Symbol] crosses. Inherited
    # @option options [Integer] tickLblSkip
    # @option options [Integer] tickMarkSkip
    def initialize(axId, crossAx, options={})
      super(axId, crossAx, options)
    end 

    def tickLblSkip=(v) Axlsx::validate_unsigned_int(v); @tickLblSkip = v; end
    def tickMarkSkip=(v) Axlsx::validate_unsigned_int(v); @tickMarkSkip = v; end

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
