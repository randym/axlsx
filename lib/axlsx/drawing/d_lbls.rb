module Axlsx
  # There are more elements in the dLbls spec that allow for
  # customizations and formatting. For now, I am just implementing the
  # basics.
  # 
  #<c:dLbls>
  #<c:dLblPos val="outEnd"/>
  #<c:showLegendKey val="0"/>
  #<c:showVal val="0"/>
  #<c:showCatName val="1"/>
  #<c:showSerName val="0"/>
  #<c:showPercent val="1"/>
  #<c:showBubbleSize val="0"/>
  #<c:showLeaderLines val="1"/>
  #</c:dLbls>

  #The DLbls class manages serialization of data labels
  class DLbls

    # These attributes are all boolean so I'm doing a bit of a hand
    # waving magic show to set up the attriubte accessors
    # @note 
    #   not all charts support all methods! 
    #   Bar3DChart and Line3DChart and ScatterChart do not support d_lbl_pos or show_leader_lines
    #    
    BOOLEAN_ATTRIBUTES = [:show_legend_key, :show_val, :show_cat_name, :show_ser_name, :show_percent, :show_bubble_size, :show_leader_lines]
    
    # creates a new DLbls object
    def initialize(options={})
      @d_lbl_pos = :bestFit
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

    # The position of the data labels in the chart
    # @see d_lbl_pos= for a list of allowed values
    # @return [Symbol]
    attr_reader :d_lbl_pos

    # @see DLbls#d_lbl_pos
    # Assigns the label postion for this data labels on this chart.
    # Allowed positions are :bestFilt, :b, :ctr, :inBase, :inEnd, :l,
    # :outEnd, :r and :t
    # The default is :best_fit
    # @param [Symbol] label_postion the postion you want to use.
    def d_lbl_pos=(label_position)
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
      str << '<c:dLbls>'
      instance_values.each { |name, value| str << "<c:#{Axlsx::camel(name, false)} val='#{value}' />" }
      str << '</c:dLbls>'
    end
  end
end

