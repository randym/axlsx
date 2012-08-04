module Axlsx
  # There are more elements in the dLbls spec that allow for
  # customizations and formatting. For now, I am just implementing the
  # basics.

  #The DLbls class manages serialization of data labels
  # showLeaderLines and leaderLines are not currently implemented
  class DLbls

    # These attributes are all boolean so I'm doing a bit of a hand
    # waving magic show to set up the attriubte accessors
    # @note 
    #   not all charts support all methods! 
    #   Bar3DChart and Line3DChart and ScatterChart do not support d_lbl_pos or show_leader_lines
    #    
    BOOLEAN_ATTRIBUTES = [:show_legend_key, :show_val, :show_cat_name, :show_ser_name, :show_percent, :show_bubble_size, :show_leader_lines]

    # creates a new DLbls object
    def initialize(chart_type, options={})
      raise ArgumentError, 'chart_type must inherit from Chart' unless chart_type.superclass == Chart
      @chart_type = chart_type
      initialize_defaults
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
    end

    # Initialize all the values to false as Excel requires them to
    # explicitly be disabled or all will show.
    def initialize_defaults
      BOOLEAN_ATTRIBUTES.each do |attr|
        self.send("#{attr}=", false)
      end
    end

    # The chart type that is using this data lables instance. 
    # This affects the xml output as not all chart types support the
    # same data label attributes. 
    attr_reader :chart_type

    # The position of the data labels in the chart
    # @see d_lbl_pos= for a list of allowed values
    # @return [Symbol]
    def d_lbl_pos
      return unless @chart_type == Pie3DChart
      @d_lbl_pos ||= :bestFit
    end

    # @see DLbls#d_lbl_pos
    # Assigns the label postion for this data labels on this chart.
    # Allowed positions are :bestFilt, :b, :ctr, :inBase, :inEnd, :l,
    # :outEnd, :r and :t
    # The default is :bestFit
    # @param [Symbol] label_position the postion you want to use.
    def d_lbl_pos=(label_position)
      return unless @chart_type == Pie3DChart
      Axlsx::RestrictionValidator.validate 'DLbls#d_lbl_pos',  [:bestFit, :b, :ctr, :inBase, :inEnd, :l, :outEnd, :r, :t], label_position
      @d_lbl_pos = label_position
    end

    # Dynamically create accessors for boolean attriubtes
    BOOLEAN_ATTRIBUTES.each do |attr|    
      class_eval %{
      # The #{attr} attribute reader
      # @return [Boolean]
      attr_reader :#{attr} 

      # The #{attr} writer
      # @param [Boolean] value The value to assign to #{attr}
      # @return [Boolean]
      def #{attr}=(value)
        Axlsx::validate_boolean(value)
        @#{attr} = value
      end
      }
    end


    # serializes the data labels
    # @return [String]
    def to_xml_string(str = '')
      validate_attributes_for_chart_type
      str << '<c:dLbls>'
      %w(d_lbl_pos show_legend_key show_val show_cat_name show_ser_name show_percent show_bubble_size show_leader_lines).each do |key|
        next unless instance_values.keys.include?(key) && instance_values[key] != nil
        str <<  "<c:#{Axlsx::camel(key, false)} val='#{instance_values[key]}' />" 
      end
      str << '</c:dLbls>'
    end

    # nills out d_lbl_pos and show_leader_lines as these attributes, while valid in the spec actually chrash excel for any chart type other than pie charts.
    def validate_attributes_for_chart_type
      return if @chart_type == Pie3DChart
      @d_lbl_pos = nil
      @show_leader_lines = nil
    end


  end
end
