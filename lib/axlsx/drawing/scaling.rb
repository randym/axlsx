# encoding: UTF-8
# frozen_string_literal: true
module Axlsx
  # The Scaling class defines axis scaling
  class Scaling

    include Axlsx::OptionsParser

    # creates a new Scaling object
    # @option options [Integer] logBase
    # @option options [Symbol] orientation
    # @option options [Float] max
    # @option options [Float] min
    def initialize(options={})
      @orientation = :minMax
      @logBase, @min, @max = nil, nil, nil
      parse_options options
    end

    # logarithmic base for a logarithmic axis.
    # must be between 2 and 1000
    # @return [Integer]
    attr_reader :logBase

    # the orientation of the axis
    # must be one of [:minMax, :maxMin]
    # @return [Symbol]
    attr_reader :orientation

    # the maximum scaling
    # @return [Float]
    attr_reader :max

    # the minimu scaling
    # @return [Float]
    attr_reader :min

    # @see logBase
    def logBase=(v) DataTypeValidator.validate "Scaling.logBase", [Integer], v, lambda { |arg| arg >= 2 && arg <= 1000}; @logBase = v; end
    # @see orientation
    def orientation=(v) RestrictionValidator.validate "Scaling.orientation", [:minMax, :maxMin], v; @orientation = v; end
    # @see max
    def max=(v) DataTypeValidator.validate "Scaling.max", Float, v; @max = v; end

    # @see min
    def min=(v) DataTypeValidator.validate "Scaling.min", Float, v; @min = v; end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = String.new)
      str << '<c:scaling>'
      str << "<c:logBase val=\"#{@logBase}\"/>" unless @logBase.nil?
      str << "<c:orientation val=\"#{@orientation}\"/>" unless @orientation.nil?
      str << "<c:min val=\"#{@min}\"/>" unless @min.nil?
      str << "<c:max val=\"#{@max}\"/>" unless @max.nil?
      str << '</c:scaling>'
    end

  end
end
