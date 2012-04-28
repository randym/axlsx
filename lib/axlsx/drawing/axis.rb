# encoding: UTF-8
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

    # specifies how the degree of label rotation
    # @return [Integer]
    attr_reader :label_rotation

    # specifies if gridlines should be shown in the chart
    # @return [Boolean]
    attr_reader :gridlines

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
      @label_rotation = 0
      @scaling = Scaling.new(:orientation=>:minMax)
      self.axPos = :b
      self.tickLblPos = :nextTo
      self.format_code = "General"
      self.crosses = :autoZero
      self.gridlines = true
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

    # Specify if gridlines should be shown for this axis
    # default true
    def gridlines=(v) Axlsx::validate_boolean(v); @gridlines = v; end

    # specifies how the perpendicular axis is crossed
    # must be one of [:autoZero, :min, :max]
    def crosses=(v) RestrictionValidator.validate "#{self.class}.crosses", [:autoZero, :min, :max], v; @crosses = v; end

    # Specify the degree of label rotation to apply to labels
    # default true
    def label_rotation=(v)
      Axlsx::validate_int(v)
      adjusted = v.to_i * 60000
      Axlsx::validate_angle(adjusted)
      @label_rotation = adjusted
    end


    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<c:axId val="' << @axId.to_s << '"/>'
      @scaling.to_xml_string str
      str << '<c:delete val="0"/>'
      str << '<c:axPos val="' << @axPos.to_s << '"/>'
      str << '<c:majorGridlines>'
      if self.gridlines == false
        str << '<c:spPr>'
        str << '<a:ln>'
        str << '<a:noFill/>'
        str << '</a:ln>'
        str << '</c:spPr>'
      end
      str << '</c:majorGridlines>'
      str << '<c:numFmt formatCode="' << @format_code << '" sourceLinked="1"/>'
      str << '<c:majorTickMark val="none"/>'
      str << '<c:minorTickMark val="none"/>'
      str << '<c:tickLblPos val="' << @tickLblPos.to_s << '"/>'
      # some potential value in implementing this in full. Very detailed!
      str << '<c:txPr><a:bodyPr rot="' << @label_rotation.to_s << '"/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr/></a:p></c:txPr>'
      str << '<c:crossAx val="' << @crossAx.to_s << '"/>'
      str << '<c:crosses val="' << @crosses.to_s << '"/>'
    end

  end
end
