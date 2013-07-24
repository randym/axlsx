# encoding: UTF-8
module Axlsx
  # the ValAxis class defines a chart value axis.
  class ValAxis < Axis

    # This element specifies how the value axis crosses the category axis.
    # must be one of [:between, :midCat]
    # @return [Symbol]
    attr_reader :cross_between
    alias :crossBetween :cross_between

    # Creates a new ValAxis object
    # @option options [Symbol] crosses_between
    def initialize(options={})
      self.cross_between = :between
      super(options)
    end

    # @see cross_between
    def cross_between=(v)
      RestrictionValidator.validate "ValAxis.cross_between", [:between, :midCat], v
      @cross_between = v
    end
    alias :crossBetween= :cross_between=

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<c:valAx>'
      super(str)
      str << '<c:crossBetween val="' << @cross_between.to_s << '"/>'
      str << '</c:valAx>'
    end

  end
end
