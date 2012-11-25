module Axlsx
  # Conditional Format Rule color scale object
  # Describes a gradated color scale in this conditional formatting rule.

  # @note The recommended way to manage these rules is via Worksheet#add_conditional_formatting
  # @see Worksheet#add_conditional_formatting
  # @see ConditionalFormattingRule#initialize
  class ColorScale

    class << self

      # These are the default conditional formatting value objects
      # that define a two tone color gradient.
      def default_cfvos
        [{:type => :min, :val => 0, :color => 'FFFF7128'},
         {:type => :max, :val => 0, :color => 'FFFFEF9C'}]
      end

      # A builder for two tone color gradient
      # @example
      #   # this creates a two tone color scale
      #   color_scale = Axlsx::ColorScale.two_tone
      # @see examples/example.rb conditional formatting examples.
      def two_tone
        self.new
      end

      # A builder for three tone color gradient
      # @example
      #   #this creates a three tone color scale
      #   color_scale = Axlsx::ColorScale.three_tone 
      # @see examples/example.rb conditional formatting examples.
      def three_tone
        self.new({:type => :min, :val => 0, :color => 'FFF8696B'},
                 {:type => :percent, :val => '50', :color => 'FFFFEB84'},
                 {:type => :max, :val => 0, :color => 'FF63BE7B'})
      end
    end
    # A simple typed list of cfvos
    # @return [SimpleTypedList]
    # @see Cfvo
    def value_objects
      @value_objects ||= Cfvos.new
    end

    # A simple types list of colors
    # @return [SimpleTypedList]
    # @see Color
    def colors
      @colors ||= SimpleTypedList.new Color
    end

    # creates a new ColorScale object.
    # @see Cfvo
    # @see Color
    # @example
    #     color_scale = Axlsx::ColorScale.new({:type => :num, :val => 0.55, :color => 'fff7696c'})
    def initialize(*cfvos)
      initialize_default_cfvos(cfvos)
      yield self if block_given?
    end

    # adds a new cfvo / color pair to the color scale and returns a hash containing
    # a reference to the newly created cfvo and color objects so you can alter the default properties.
    # @return [Hash] a hash with :cfvo and :color keys referencing the newly added objects.
    # @param [Hash] options options for the new cfvo and color objects
    # @option [Symbol] type The type of cfvo you to add
    # @option [Any] val The value of the cfvo to add
    # @option [String] The rgb color for the cfvo
    def add(options={})
      value_objects << Cfvo.new(:type => options[:type] || :min, :val => options[:val] || 0)
      colors << Color.new(:rgb => options[:color] || "FF000000")
      {:cfvo => value_objects.last, :color => colors.last}
    end


    # removes the cfvo and color pair at the index specified.
    # @param [Integer] index The index of the cfvo and color object to delete
    # @note you cannot remove the first two cfvo and color pairs
    def delete_at(index=2)
      value_objects.delete_at index
      colors.delete_at index
    end

    # Serialize this color_scale object data to an xml string
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<colorScale>'
      value_objects.to_xml_string(str)
      colors.each { |color| color.to_xml_string(str) }
      str << '</colorScale>'
    end

    private
    # There has got to be cleaner way of merging these arrays.
    def initialize_default_cfvos(user_cfvos)
      defaults = self.class.default_cfvos
      user_cfvos.each_with_index do |cfvo, index|
        if index < defaults.size
          cfvo = defaults[index].merge(cfvo)
        end
        add cfvo
      end
      while colors.size < defaults.size
        add defaults[colors.size - 1]
      end
    end
  end
end
