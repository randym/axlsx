module Axlsx
  # Conditional Format Value Object
  # Describes the values of the interpolation points in a gradient scale. This object is used by ColorScale, DataBar and IconSet classes
  #
  # @note The recommended way to manage these rules is via Worksheet#add_conditional_formatting
  # @see Worksheet#add_conditional_formatting
  # @see ConditionalFormattingRule#initialize
  #
  class Cfvo

    # Type (ST_CfvoType)
    # The type of this conditional formatting value object. options are num, percent, max, min, formula and percentile
    # @return [Symbol]
    attr_reader :type


    # Type (xsd:boolean)
    # For icon sets, determines whether this threshold value uses the greater than or equal to operator. 0 indicates 'greater than' is used instead of 'greater than or equal to'.
    # The default value is true
    # @return [Boolean]
    attr_reader :gte


    # Type (ST_Xstring)
    # The value of the conditional formatting object
    # This library will accept any value so long as it supports to_s
    attr_reader :val

    # Creates a new Cfvo object
    # @option options [Symbol] type The type of conditional formatting value object
    # @option options [Boolean]  gte threshold value usage indicator
    # @option options [String] val The value of the conditional formatting object
    def initialize(options={})
      @gte = true
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
    end

    # @see type
    def type=(v); Axlsx::validate_conditional_formatting_value_object_type(v); @type = v end

    # @see gte
    def gte=(v); Axlsx::validate_boolean(v); @gte = v end

    # @see val
    def val=(v)
      raise ArgumentError, "#{v.inspect} must respond to to_s" unless v.respond_to?(:to_s)
      @val = v.to_s
    end

    # serialize the Csvo object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<cfvo '
      str << instance_values.map { |key, value| '' << key << '="' << value.to_s << '"' }.join(' ')
      str << ' />'
    end

  end
end
