module Axlsx
  #       <c:dLbls>
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
    BOOLEAN_ATTRIBUTES = [:show_legend_key, :show_val, :show_cat_name, :show_ser_name, :show_percent, :show_bubble_size, :show_leader_lines]
    
    # creates a new DLbls object
    def initialize(options={})
      d_lbl_pos = :best_fit
      initialize_defaults
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end

    end
    
    def initialize_defaults
      BOOLEAN_ATTRIBUTES.each do |attr|
        self.send("#{attr}=", false)
      end
    end

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
   
    BOOLEAN_ATTRIBUTES.each do |attr|    
      class_eval %{
      attr_reader :#{attr} 

      def #{attr}=(v)
        Axlsx::validate_boolean(v)
        @#{attr} = v
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

