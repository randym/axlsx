module Axlsx
  # There are more elements in the dTable spec that allow for
  # customizations and formatting. For now, I am just implementing the
  # basics.
  class DTable

    include Axlsx::Accessors
    include Axlsx::OptionsParser
    # creates a new DTable object
    def initialize(chart_type, options={})
      raise ArgumentError, 'chart_type must inherit from Chart' unless [Chart, LineChart].include?(chart_type.superclass)
      @chart_type = chart_type
      initialize_defaults
      parse_options options
    end

    # These attributes are all boolean so I'm doing a bit of a hand
    # waving magic show to set up the attriubte accessors
    # @note
    #   not all charts support all methods!
    #
    boolean_attr_accessor :show_horz_border,
                          :show_vert_border,
                          :show_outline,
                          :show_keys

    # Initialize all the values to false as Excel requires them to
    # explicitly be disabled or all will show.
    def initialize_defaults
      [:show_horz_border, :show_vert_border,
        :show_outline, :show_keys].each do |attr|
        self.send("#{attr}=", false)
      end
    end

    # The chart type that is using this data table instance.
    # This affects the xml output as not all chart types support the
    # same data table attributes.
    attr_reader :chart_type

    # serializes the data labels
    # @return [String]
    def to_xml_string(str = '')
      # validate_attributes_for_chart_type
      str << '<c:dTable>'
      %w(show_horz_border show_vert_border show_outline show_keys).each do |key|
        next unless instance_values.keys.include?(key) && instance_values[key] != nil
        str <<  "<c:#{Axlsx::camel(key, false)} val='#{instance_values[key]}' />"
      end
      str << '</c:dTable>'
    end

    # nills out d_lbl_pos and show_leader_lines as these attributes, while valid in the spec actually chrash excel for any chart type other than pie charts.
    # def validate_attributes_for_chart_type
    #   return if @chart_type == Pie3DChart
    #   @d_lbl_pos = nil
    #   @show_leader_lines = nil
    # end


  end
end
