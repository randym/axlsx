module Axlsx
  # Conditional Format Rule color scale object
  # Describes a gradated color scale in this conditional formatting rule.

  # @note The recommended way to manage these rules is via Worksheet#add_conditional_formatting
  # @see Worksheet#add_conditional_formatting
  # @see ConditionalFormattingRule#initialize
  class ColorScale

    # A simple typed list of cfvos
    # @return [SimpleTypedList]
    # @see Cfvo
    attr_reader :value_objects

    # A simple types list of colors
    # @return [SimpleTypedList]
    # @see Color
    attr_reader :colors

    # creates a new ColorScale object.
    # This method will yield it self so you can alter the properites of the defauls conditional formating value object (cfvo and colors
    # Two value objects and two colors are created on initialization and cannot be deleted.
    # @see Cfvo
    # @see Color
    def initialize
      initialize_value_objects
      initialize_colors
      yield self if block_given?
    end

    # adds a new cfvo / color pair to the color scale and returns a hash containing
    # a reference to the newly created cfvo and color objects so you can alter the default properties.
    # @return [Hash] a hash with :cfvo and :color keys referencing the newly added objects.
    def add(options={})
      @value_objects << Cfvo.new(:type => options[:type] || :min, :val => options[:val] || 0)
      @colors << Color.new(:rgb => options[:color] || "FF000000")
      {:cfvo => @value_objects.last, :color => @colors.last}
    end


    # removes the cfvo and color pair at the index specified.
    # @param [Integer] index The index of the cfvo and color object to delete
    # @note you cannot remove the first two cfvo and color pairs
    def delete_at(index=2)
      @value_objects.delete_at index
      @colors.delete_at index
    end

    # Serialize this color_scale object data to an xml string
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<colorScale>'
      @value_objects.each { |cfvo| cfvo.to_xml_string(str) }
      @colors.each { |color| color.to_xml_string(str) }
      str << '</colorScale>'
    end

    private

    # creates the initial cfvo objects
    def initialize_value_objects
      @value_objects = SimpleTypedList.new Cfvo
      @value_objects.concat [Cfvo.new(:type => :min, :val => 0), Cfvo.new(:type => :max, :val => 0)]
      @value_objects.lock
    end

    # creates the initial color objects
    def initialize_colors
      @colors = SimpleTypedList.new Color
      @colors.concat [Color.new(:rgb => "FFFF0000"), Color.new(:rgb => "FF0000FF")]
      @colors.lock
    end

  end
end
