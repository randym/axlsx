# frozen_string_literal: true
module Axlsx

  # The SheetPr class manages serialization of a worksheet's sheetPr element.
  class SheetPr
    include Axlsx::OptionsParser
    include Axlsx::Accessors
    include Axlsx::SerializedAttributes

    serializable_attributes :sync_horizontal,
                            :sync_vertical,
                            :transition_evaluation,
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
                          :transition_evaluation,
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
      @outline_pr = nil
      parse_options options
    end

    # The worksheet these properties apply to!
    # @return [Worksheet]
    attr_reader :worksheet

    # The tab color of the sheet.
    # @return [Color]
    attr_reader :tab_color

    # Serialize the object
    # @param [String] str serialized output will be appended to this object if provided.
    # @return [String]
    def to_xml_string(str = String.new)
      update_properties
      str << "<sheetPr #{serialized_attributes}>"
      tab_color.to_xml_string(str, 'tabColor') if tab_color
      outline_pr.to_xml_string(str) if @outline_pr
      page_setup_pr.to_xml_string(str)
      str << "</sheetPr>"
    end

    # The PageSetUpPr for this sheet pr object
    # @return [PageSetUpPr]
    def page_setup_pr
      @page_setup_pr ||= PageSetUpPr.new
    end
    
    # The OutlinePr for this sheet pr object
    # @return [OutlinePr]
    def outline_pr
      @outline_pr ||= OutlinePr.new
    end

    # @see tab_color
    def tab_color=(v)
      @tab_color ||= Color.new(:rgb => v)
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
