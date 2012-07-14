module Axlsx
  # Conditional Format Rule data bar object
  # Describes a data bar conditional formatting rule.

  # @note The recommended way to manage these rules is via Worksheet#add_conditional_formatting
  # @see Worksheet#add_conditional_formatting
  # @see ConditionalFormattingRule#initialize
  class DataBar

    # instance values that must be serialized as their own elements - e.g. not attributes.
    CHILD_ELEMENTS = [:value_objects, :color]

    # minLength attribute
    # The minimum length of the data bar, as a percentage of the cell width.
    # The default value is 10
    # @return [Integer]
    attr_reader :minLength

    # maxLength attribute
    # The maximum length of the data bar, as a percentage of the cell width.
    # The default value is 90
    # @return [Integer]
    attr_reader :maxLength

    # maxLength attribute
    # Indicates whether to show the values of the cells on which this data bar is applied.
    # The default value is true
    # @return [Boolean]
    attr_reader :showValue

    # A simple typed list of cfvos
    # @return [SimpleTypedList]
    # @see Cfvo
    attr_reader :value_objects

    # color
    # the color object used in the data bar formatting
    # @return [Color]
    def color
      @color ||= Color.new :rgb => "FF0000FF"
    end

    # Creates a new data bar conditional formatting object
    # @option options [Integer] minLength
    # @option options [Integer] maxLength
    # @option options [Boolean] showValue
    # @option options [String] color - the rbg value used to color the bars
    def initialize(options = {})
      @minLength = 10
      @maxLength = 90
      @showValue = true
      initialize_value_objects
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
      yield self if block_given?
    end

    # @see minLength
    def minLength=(v); Axlsx.validate_unsigned_int(v); @minLength = v end
    # @see maxLength
    def maxLength=(v); Axlsx.validate_unsigned_int(v); @maxLength = v end

    # @see showValue
    def showValue=(v); Axlsx.validate_boolean(v); @showValue = v end

    # Sets the color for the data bars.
    # @param [Color|String] v The color object, or rgb string value to apply
    def color=(v)
      @color = v if v.is_a? Color
      self.color.rgb = v if v.is_a? String
      @color
    end

    # Serialize this object to an xml string
    # @param [String] str
    # @return [String]
    def to_xml_string(str="")
      str << '<dataBar '
      str << instance_values.map { |key, value| '' << key << '="' << value.to_s << '"' unless CHILD_ELEMENTS.include?(key.to_sym) }.join(' ')
      str << '>'
      @value_objects.each { |cfvo| cfvo.to_xml_string(str) }
      self.color.to_xml_string(str)
      str << '</dataBar>'
    end

    private

    # Initalize the simple typed list of value objects
    # I am keeping this private for now as I am not sure what impact changes to the required two cfvo objects will do.
    def initialize_value_objects
      @value_objects = SimpleTypedList.new Cfvo
      @value_objects.concat [Cfvo.new(:type => :min, :val => 0), Cfvo.new(:type => :max, :val => 0)]
      @value_objects.lock
    end
  end
end
