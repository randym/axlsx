module Axlsx
  # the ValAxis class defines a chart value axis.
  class ValAxis < Axis

    # This element specifies whether the value axis crosses the category axis between categories.
    # must be one of [:between, :midCat]
    # @return [Symbol]
    attr_accessor :crossBetween

    # Creates a new ValAxis object
    # @param [Integer] axId the id of this axis
    # @param [Integer] crossAx the id of the perpendicular axis
    # @option options [Symbol] axPos
    # @option options [Symbol] crosses
    # @option options [Symbol] tickLblPos
    # @option options [Symbol] crossesBetween
    def initialize(axId, crossAx, options={})
      @crossBetween = :between
      super(axId, crossAx, options)
    end
    
    def crossBetween=(v) RestrictionValidator.validate "ValAxis.crossBetween", [:between, :midCat], v; @crossBetween = v; end

    # Serializes the value  axis
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml.send('c:valAx') {
        super(xml)
        xml.send('c:crossBetween', :val=>@crossBetween)
      }
    end
  end
end
