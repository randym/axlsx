module Axlsx

  # The SheetPr class manages serialization fo a worksheet's sheetPr element.
  class SheetPr
    include Axlsx::OptionsParser
    include Axlsx::Accessors
    include Axlsx::SerializedAttributes

    serializable_attributes :sync_horizontal,
                            :sync_vertical,
                            :transtion_evaluation,
                            :transition_entry,
                            :published,
                            :filter_mode,
                            :enable_format_conditions_calculation,
                            :code_name,
                            :sync_ref

    # These attributes are all boolean so I'm doing a bit of a hand
    # waving magic show to set up the attriubte accessors
    boolean_attr_accessor :sync_horizontal,
                          :sync_vertical,
                          :transtion_evaluation,
                          :transition_entry,
                          :published,
                          :filter_mode,
                          :enable_format_conditions_calculation

    string_attr_accessor :code_name, :sync_ref

    # Creates a new SheetPr object
    # @param [Worksheet] worksheet The worksheet that owns this SheetPr object
    def initialize(worksheet, options={})
      raise ArgumentError, "you must provide a worksheet" unless worksheet.is_a?(Worksheet)
      @worksheet = worksheet
      parse_options options
    end

    # The worksheet these properties apply to!
    # @return [Worksheet]
    attr_reader :worksheet

    # Serialize the object
    # @param [String] str serialized output will be appended to this object if provided.
    # @return [String]
    def to_xml_string(str = '')
      update_properties
      str << "<sheetPr #{serialized_attributes}>"
      page_setup_pr.to_xml_string(str)
      str << "</sheetPr>"
    end

    # The PageSetUpPr for this sheet pr object
    # @return [PageSetUpPr]
    def page_setup_pr
      @page_setup_pr ||= PageSetUpPr.new
    end

    private

    def update_properties
      page_setup_pr.fit_to_page = worksheet.fit_to_page?
      if worksheet.auto_filter.columns.size > 0
        self.filter_mode = 1
        self.enable_format_conditions_calculation = 1
      end
    end
  end
end
