module Axlsx
 # the access class defines common properties and values for chart axis
  class Axis


    # the id of the axis
    # @return [Integer]
    attr_reader :axId

    # The perpendicular axis
    # @return [Integer]
    attr_reader :crossAx

    # The scaling of the axis
    # @return [Scaling]
    attr_reader :scaling
    
    # The position of the axis
    # must be one of [:l, :r, :t, :b]
    # @return [Symbol]
    attr_accessor :axPos

    # the position of the tick labels
    # must be one of [:nextTo, :high, :low]
    # @return [Symbol]
    attr_accessor :tickLblPos


    # The number format format code for this axis
    # @return [String]
    attr_accessor :format_code

    # specifies how the perpendicular axis is crossed
    # must be one of [:autoZero, :min, :max]
    # @return [Symbol]
    attr_accessor :crosses 

    # Creates an Axis object
    # @param [Integer] axId the id of this axis
    # @param [Integer] crossAx the id of the perpendicular axis
    # @option options [Symbol] axPos
    # @option options [Symbol] crosses
    # @option options [Symbol] tickLblPos
    def initialize(axId, crossAx, options={})
      Axlsx::validate_unsigned_int(axId)
      Axlsx::validate_unsigned_int(crossAx)
      @axId = axId
      @crossAx = crossAx
      self.axPos = :l
      self.tickLblPos = :nextTo
      @scaling = Scaling.new(:orientation=>:minMax)
      @formatCode = ""
      self.crosses = :autoZero
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
    end

    def axPos=(v) RestrictionValidator.validate "#{self.class}.axPos", [:l, :r, :b, :t], v; @axPos = v; end
    def tickLblPos=(v) RestrictionValidator.validate "#{self.class}.tickLblPos", [:nextTo, :high, :low], v; @tickLblPos = v; end
    def format_code=(v) Axlsx::validate_string(v); @formatCode = v; end
    def crosses=(v) RestrictionValidator.validate "#{self.class}.crosses", [:autoZero, :min, :max], v; @crosses = v; end

    # Serializes the common axis
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml.send('c:axId', :val=>@axId)
      @scaling.to_xml(xml)
      xml.send('c:axPos', :val=>@axPos)
      xml.send('c:majorGridlines')
      xml.send('c:numFmt', :formatCode => @format_code, :sourceLinked=>"1")
      xml.send('c:tickLblPos', :val=>@tickLblPos)
      xml.send('c:crossAx', :val=>@crossAx)
      xml.send('c:crosses', :val=>@crosses)
    end    
  end
end
