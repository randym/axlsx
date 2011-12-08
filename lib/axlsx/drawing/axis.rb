module Axlsx
 # the access class defines common properties and values for a chart axis.
  class Axis

    # the id of the axis. 
    # @return [Integer]
    attr_reader :axId

    # The perpendicular axis
    # @return [Integer]
    attr_reader :crossAx

    # The scaling of the axis
    # @see Scaling
    # @return [Scaling]
    attr_reader :scaling
    
    # The position of the axis
    # must be one of [:l, :r, :t, :b]
    # @return [Symbol]
    attr_reader :axPos

    # the position of the tick labels
    # must be one of [:nextTo, :high, :low]
    # @return [Symbol]
    attr_reader :tickLblPos

    # The number format format code for this axis
    # default :General
    # @return [String]
    attr_reader :format_code

    # specifies how the perpendicular axis is crossed
    # must be one of [:autoZero, :min, :max]
    # @return [Symbol]
    attr_reader :crosses 

    # Creates an Axis object
    # @param [Integer] axId the id of this axis
    # @param [Integer] crossAx the id of the perpendicular axis
    # @option options [Symbol] axPos
    # @option options [Symbol] crosses
    # @option options [Symbol] tickLblPos
    # @raise [ArgumentError] If axId or crossAx are not unsigned integers
    def initialize(axId, crossAx, options={})
      Axlsx::validate_unsigned_int(axId)
      Axlsx::validate_unsigned_int(crossAx)
      @axId = axId
      @crossAx = crossAx
      @format_code = "General"
      @scaling = Scaling.new(:orientation=>:minMax)      
      self.axPos = :b
      self.tickLblPos = :nextTo
      self.format_code = "General"
      self.crosses = :autoZero
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
    end
    # The position of the axis
    # must be one of [:l, :r, :t, :b]
    def axPos=(v) RestrictionValidator.validate "#{self.class}.axPos", [:l, :r, :b, :t], v; @axPos = v; end

    # the position of the tick labels
    # must be one of [:nextTo, :high, :low1]
    def tickLblPos=(v) RestrictionValidator.validate "#{self.class}.tickLblPos", [:nextTo, :high, :low], v; @tickLblPos = v; end

    # The number format format code for this axis
    # default :General
    def format_code=(v) Axlsx::validate_string(v); @format_code = v; end

    # specifies how the perpendicular axis is crossed
    # must be one of [:autoZero, :min, :max]
    def crosses=(v) RestrictionValidator.validate "#{self.class}.crosses", [:autoZero, :min, :max], v; @crosses = v; end

    # Serializes the common axis
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml.axId :val=>@axId
      @scaling.to_xml(xml)
      xml.delete :val=>0
      xml.axPos :val=>@axPos
      xml.majorGridlines
      xml.numFmt :formatCode => @format_code, :sourceLinked=>"1"
      xml.majorTickMark :val=>"none"
      xml.minorTickMark :val=>"none"
      xml.tickLblPos :val=>@tickLblPos
      xml.crossAx :val=>@crossAx
      xml.crosses :val=>@crosses
    end    
  end
end
