# encoding: UTF-8
module Axlsx
  #A CatAxis object defines a chart category axis
  class CatAxis < Axis

    # From the docs: This element specifies that this axis is a date or text axis based on the data that is used for the axis labels, not a specific choice.
    # @return [Boolean]
    attr_reader :auto

    # specifies how the perpendicular axis is crossed
    # must be one of [:ctr, :l, :r]
    # @return [Symbol]
    attr_reader :lblAlgn

    # The offset of the labels
    # must be between a string between 0 and 1000
    # @return [Integer]
    attr_reader :lblOffset

    # regex for validating label offset
    LBL_OFFSET_REGEX = /0*(([0-9])|([1-9][0-9])|([1-9][0-9][0-9])|1000)%/

    # Creates a new CatAxis object
    # @param [Integer] axId the id of this axis. Inherited
    # @param [Integer] crossAx the id of the perpendicular axis. Inherited
    # @option options [Symbol] axPos. Inherited
    # @option options [Symbol] tickLblPos. Inherited
    # @option options [Symbol] crosses. Inherited
    # @option options [Boolean] auto
    # @option options [Symbol] lblAlgn
    # @option options [Integer] lblOffset
    def initialize(axId, crossAx, options={})
      self.auto = 1
      self.lblAlgn = :ctr
      self.lblOffset = "100%"
      super(axId, crossAx, options)
    end

    # From the docs: This element specifies that this axis is a date or text axis based on the data that is used for the axis labels, not a specific choice.
    def auto=(v) Axlsx::validate_boolean(v); @auto = v; end

    # specifies how the perpendicular axis is crossed
    # must be one of [:ctr, :l, :r]
    def lblAlgn=(v) RestrictionValidator.validate "#{self.class}.lblAlgn", [:ctr, :l, :r], v; @lblAlgn = v; end

    # The offset of the labels
    # must be between a string between 0 and 1000
    def lblOffset=(v) RegexValidator.validate "#{self.class}.lblOffset", LBL_OFFSET_REGEX, v; @lblOffset = v; end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<c:catAx>'
      super(str)
      str << '<c:auto val="' << @auto.to_s << '"/>'
      str << '<c:lblAlgn val="' << @lblAlgn.to_s << '"/>'
      str << '<c:lblOffset val="' << @lblOffset.to_s << '"/>'
      str << '</c:catAx>'
    end

  end


end
