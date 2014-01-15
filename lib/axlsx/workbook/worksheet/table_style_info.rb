module Axlsx

  # The table style info class manages style attributes for defined tables in
  # a worksheet
  class TableStyleInfo
    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes
    include Axlsx::Accessors
    # creates a new TableStyleInfo instance
    # @param [Hash] options
    # @option [Boolean] show_first_column indicates if the first column should
    #                   be shown
    # @option [Boolean] show_last_column indicates if the last column should
    #                   be shown
    # @option [Boolean] show_column_stripes indicates if column stripes should
    #                   be shown
    # @option [Boolean] show_row_stripes indicates if row stripes should be shown
    # @option [String] name The name of the style to apply to your table. 
    #                  Only predefined styles are currently supported.
    #                  @see Annex G. (normative) Predefined SpreadsheetML Style Definitions in part 1 of the specification.
    def initialize(options = {})
      initialize_defaults
      @name = 'TableStyleMedium9'
      parse_options options
    end

    # boolean attributes for this object
    boolean_attr_accessor :show_first_column, :show_last_column, :show_row_stripes, :show_column_stripes
    serializable_attributes :show_first_column, :show_last_column, :show_row_stripes, :show_column_stripes,
      :name

    # Initialize all the values to false as Excel requires them to
    # explicitly be disabled or all will show.
    def initialize_defaults
      %w(show_first_column show_last_column show_row_stripes show_column_stripes).each do |attr|
        self.send("#{attr}=", 0)
      end
    end

    # The name of the table style.
    attr_accessor :name

    # seralizes this object to an xml string
    # @param [String] str the string to contact this objects serialization to.
    def to_xml_string(str = '')
      serialized_tag('tableStyleInfo', str)
    end
  end
end
