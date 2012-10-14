# encoding: UTF-8
module Axlsx
  # A GradientFill defines the color and positioning for gradiant cell fill.
  # @see Open Office XML Part 1 §18.8.24
  class GradientFill

    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes

    # Creates a new GradientFill object
    # @option options [Symbol] type
    # @option options [Float] degree
    # @option options [Float] left
    # @option options [Float] right
    # @option options [Float] top
    # @option options [Float] bottom
    def initialize(options={})
      options[:type] ||= :linear
      parse_options options
      @stop = SimpleTypedList.new GradientStop
    end

    serializable_attributes :type, :degree, :left, :right, :top, :bottom

    # The type of gradient.
    # @note
    #  valid options are
    #   :linear
    #   :path
    # @return [Symbol]
    attr_reader :type

    # Angle of the linear gradient
    # @return [Float]
    attr_reader :degree

    # Percentage format left
    # @return [Float]
    attr_reader :left

    # Percentage format right
    # @return [Float]
    attr_reader :right

    # Percentage format top
    # @return [Float]
    attr_reader :top

    # Percentage format bottom
    # @return [Float]
    attr_reader :bottom

    # Collection of stop objects
    # @return [SimpleTypedList]
    attr_reader :stop

      # @see type
    def type=(v) Axlsx::validate_gradient_type v; @type = v end
    # @see degree
    def degree=(v) Axlsx::validate_float v; @degree = v end
    # @see left
    def left=(v) DataTypeValidator.validate "GradientFill.left", Float, v, lambda { |arg| arg >= 0.0 && arg <= 1.0}; @left = v end
    # @see right
    def right=(v) DataTypeValidator.validate "GradientFill.right", Float, v, lambda { |arg| arg >= 0.0 && arg <= 1.0}; @right = v end
    # @see top
    def top=(v) DataTypeValidator.validate "GradientFill.top", Float, v, lambda { |arg| arg >= 0.0 && arg <= 1.0}; @top = v end
    # @see bottom
    def bottom=(v) DataTypeValidator.validate "GradientFill.bottom", Float, v, lambda { |arg| arg >= 0.0 && arg <= 1.0}; @bottom= v end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<gradientFill '
      serialized_attributes str
      str << '>'
      @stop.each { |s| s.to_xml_string(str) }
      str << '</gradientFill>'
    end
  end
end
