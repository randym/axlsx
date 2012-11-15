module Axlsx
  # Conditional Format Rule data bar object
  # Describes a data bar conditional formatting rule.

  # @note The recommended way to manage these rules is via Worksheet#add_conditional_formatting
  # @see Worksheet#add_conditional_formatting
  # @see ConditionalFormattingRule#initialize
  class DataBar

    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes

    class << self
      # This differs from ColorScale. There must be exactly two cfvos one color
      def default_cfvos
        [{:type => :min, :val => "0"},
         {:type => :max, :val => "0"}]
      end
    end

    # Creates a new data bar conditional formatting object
    # @param [Hash] options
    # @option options [Integer] minLength
    # @option options [Integer] maxLength
    # @option options [Boolean] showValue
    # @option options [String] color - the rbg value used to color the bars
    # @param [Array] cfvos hashes defining the gradient interpolation points for this formatting.
    def initialize(options = {}, *cfvos)
      @min_length = 10
      @max_length = 90
      @show_value = true
      parse_options options
      initialize_cfvos(cfvos)
      yield self if block_given?
    end

    serializable_attributes :min_length, :max_length, :show_value

    # instance values that must be serialized as their own elements - e.g. not attributes.
    CHILD_ELEMENTS = [:value_objects, :color]

    # minLength attribute
    # The minimum length of the data bar, as a percentage of the cell width.
    # The default value is 10
    # @return [Integer]
    attr_reader :min_length
    alias :minLength :min_length

    # maxLength attribute
    # The maximum length of the data bar, as a percentage of the cell width.
    # The default value is 90
    # @return [Integer]
    attr_reader :max_length
    alias :maxLength :max_length

    # maxLength attribute
    # Indicates whether to show the values of the cells on which this data bar is applied.
    # The default value is true
    # @return [Boolean]
    attr_reader :show_value
    alias :showValue :show_value

    # A simple typed list of cfvos
    # @return [SimpleTypedList]
    # @see Cfvo
    def value_objects
      @value_objects ||= Cfvos.new
    end

    # color
    # the color object used in the data bar formatting
    # @return [Color]
    def color
      @color ||= Color.new :rgb => "FF0000FF"
    end

    # @see minLength
    def min_length=(v)
      Axlsx.validate_unsigned_int(v)
      @min_length = v
    end
    alias :minLength= :min_length=

      # @see maxLength
      def max_length=(v)
        Axlsx.validate_unsigned_int(v)
        @max_length = v
      end
    alias :maxLength= :max_length=

      # @see showValue
      def show_value=(v)
        Axlsx.validate_boolean(v)
        @show_value = v
      end
    alias :showValue= :show_value=

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
      serialized_attributes str
      str << '>'
      value_objects.to_xml_string(str)
      self.color.to_xml_string(str)
      str << '</dataBar>'
    end

    private

    def initialize_cfvos(cfvos)
      self.class.default_cfvos.each_with_index.map do |default, index|
        if index < cfvos.size
          value_objects << Cfvo.new(default.merge(cfvos[index]))
        else
          value_objects << Cfvo.new(default)
        end
      end
    end

  end
end
