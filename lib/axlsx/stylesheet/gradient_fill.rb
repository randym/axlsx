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
    def left=(v)
      validate_format_percentage "GradientFill.left", v
      @left = v
    end

    # @see right
    def right=(v)
     validate_format_percentage "GradientFill.right", v
     @right = v
    end

    # @see top
    def top=(v)
      validate_format_percentage "GradientFill.top", v
      @top = v
    end

    # @see bottom
    def bottom=(v)
      validate_format_percentage "GradientFill.bottom", v
      @bottom = v
    end

    # validates that the value provided is between 0.0 and 1.0
    def validate_format_percentage(name, value)
      DataTypeValidator.validate name, Float, value, lambda { |arg| arg >= 0.0 && arg <= 1.0}
    end

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
