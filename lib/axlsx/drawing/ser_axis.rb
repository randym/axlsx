# encoding: UTF-8
module Axlsx
  #A SerAxis object defines a series axis
  class SerAxis < Axis

    # The number of tick lables to skip between labels
    # @return [Integer]
    attr_reader :tickLblSkip

    # The number of tickmarks to be skipped before the next one is rendered.
    # @return [Boolean]
    attr_reader :tickMarkSkip

    # Creates a new SerAxis object
    # @param [Integer] axId the id of this axis. Inherited
    # @param [Integer] crossAx the id of the perpendicular axis. Inherited
    # @option options [Symbol] axPos. Inherited
    # @option options [Symbol] tickLblPos. Inherited
    # @option options [Symbol] crosses. Inherited
    # @option options [Integer] tickLblSkip
    # @option options [Integer] tickMarkSkip
    def initialize(axId, crossAx, options={})
      @tickLblSkip, @tickMarkSkip = 1, 1
      super(axId, crossAx, options)
    end

    # @see tickLblSkip
    def tickLblSkip=(v) Axlsx::validate_unsigned_int(v); @tickLblSkip = v; end

    # @see tickMarkSkip
    def tickMarkSkip=(v) Axlsx::validate_unsigned_int(v); @tickMarkSkip = v; end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<c:serAx>'
      super(str)
      str << '<c:tickLblSkip val="' << @tickLblSkip.to_s << '"/>' unless @tickLblSkip.nil?
      str << '<c:tickMarkSkip val="' << @tickMarkSkip.to_s << '"/>' unless @tickMarkSkip.nil?
      str << '</c:serAx>'
    end
  end


end
