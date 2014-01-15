# encoding: UTF-8
module Axlsx
  #A CatAxis object defines a chart category axis
  class CatAxis < Axis

    # Creates a new CatAxis object
    # @option options [Integer] tick_lbl_skip
    # @option options [Integer] tick_mark_skip
    def initialize(options={})
      @tick_lbl_skip = 1
      @tick_mark_skip = 1
      self.auto = 1
      self.lbl_algn = :ctr
      self.lbl_offset = "100"
      super(options)
    end

    # From the docs: This element specifies that this axis is a date or text axis based on the data that is used for the axis labels, not a specific choice.
    # @return [Boolean]
    attr_reader :auto

    # specifies how the perpendicular axis is crossed
    # must be one of [:ctr, :l, :r]
    # @return [Symbol]
    attr_reader :lbl_algn
    alias :lblAlgn :lbl_algn

    # The offset of the labels
    # must be between a string between 0 and 1000
    # @return [Integer]
    attr_reader :lbl_offset
    alias :lblOffset :lbl_offset

    # The number of tick lables to skip between labels
    # @return [Integer]
    attr_reader :tick_lbl_skip
    alias :tickLblSkip :tick_lbl_skip

    # The number of tickmarks to be skipped before the next one is rendered.
    # @return [Boolean]
    attr_reader :tick_mark_skip
    alias :tickMarkSkip :tick_mark_skip

    # regex for validating label offset
    LBL_OFFSET_REGEX = /0*(([0-9])|([1-9][0-9])|([1-9][0-9][0-9])|1000)/

    # @see tick_lbl_skip
    def tick_lbl_skip=(v) Axlsx::validate_unsigned_int(v); @tick_lbl_skip = v; end
    alias :tickLblSkip= :tick_lbl_skip=

    # @see tick_mark_skip
    def tick_mark_skip=(v) Axlsx::validate_unsigned_int(v); @tick_mark_skip = v; end
    alias :tickMarkSkip= :tick_mark_skip=

    # From the docs: This element specifies that this axis is a date or text axis based on the data that is used for the axis labels, not a specific choice.
    def auto=(v) Axlsx::validate_boolean(v); @auto = v; end

    # specifies how the perpendicular axis is crossed
    # must be one of [:ctr, :l, :r]
    def lbl_algn=(v) RestrictionValidator.validate "#{self.class}.lbl_algn", [:ctr, :l, :r], v; @lbl_algn = v; end
    alias :lblAlgn= :lbl_algn=

    # The offset of the labels
    # must be between a string between 0 and 1000
    def lbl_offset=(v) RegexValidator.validate "#{self.class}.lbl_offset", LBL_OFFSET_REGEX, v; @lbl_offset = v; end
    alias :lblOffset= :lbl_offset=

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<c:catAx>'
      super(str)
      str << ('<c:auto val="' << @auto.to_s << '"/>')
      str << ('<c:lblAlgn val="' << @lbl_algn.to_s << '"/>')
      str << ('<c:lblOffset val="' << @lbl_offset.to_i.to_s << '"/>')
      str << ('<c:tickLblSkip val="' << @tick_lbl_skip.to_s << '"/>')
      str << ('<c:tickMarkSkip val="' << @tick_mark_skip.to_s << '"/>')
      str << '</c:catAx>'
    end

  end


end
