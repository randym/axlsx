# encoding: UTF-8
module Axlsx
  # Data validation allows the validation of cell data
  #
  # @note The recommended way to manage data validations is via Worksheet#add_data_validation
  # @see Worksheet#add_data_validation
  class DataValidation
    include Axlsx::OptionsParser

    # Creates a new {DataValidation} object
    # @option options [String] formula1
    # @option options [String] formula2
    # @option options [Boolean] allowBlank - A boolean value indicating whether the data validation allows the use of empty or blank entries.
    # @option options [String] error - Message text of error alert.
    # @option options [Symbol] errorStyle - The style of error alert used for this data validation.
    # @option options [String] errorTitle - itle bar text of error alert.
    # @option options [Symbol] operator - The relational operator used with this data validation.
    # @option options [String] prompt - Message text of input prompt.
    # @option options [String] promptTitle - Title bar text of input prompt.
    # @option options [Boolean] showDropDown - A boolean value indicating whether to display a dropdown combo box for a list type data validation
    # @option options [Boolean] showErrorMessage - A boolean value indicating whether to display the error alert message when an invalid value has been entered, according to the criteria specified.
    # @option options [Boolean] showInputMessage - A boolean value indicating whether to display the input prompt message.
    # @option options [String] sqref - Range over which data validation is applied, in "A1:B2" format.
    # @option options [Symbol] type - The type of data validation.
    def initialize(options={})
      # defaults
      @formula1 = @formula2 = @error = @errorTitle = @operator = @prompt = @promptTitle = @sqref = nil
      @allowBlank = @showErrorMessage = true
      @showDropDown = @showInputMessage = false
      @type = :none
      @errorStyle = :stop
      parse_options options
    end

    # instance values that must be serialized as their own elements - e.g. not attributes.
    CHILD_ELEMENTS = [:formula1, :formula2]

    # Formula1
    # Available for type whole, decimal, date, time, textLength, list, custom
    # @see type
    # @return [String]
    # default nil
    attr_reader :formula1

    # Formula2
    # Available for type whole, decimal, date, time, textLength
    # @see type
    # @return [String]
    # default nil
    attr_reader :formula2

    # Allow Blank
    # A boolean value indicating whether the data validation allows the use of empty or blank
    # entries. 1 means empty entries are OK and do not violate the validation constraints.
    # Available for type whole, decimal, date, time, textLength, list, custom
    # @see type
    # @return [Boolean]
    # default true
    attr_reader :allowBlank

    # Error Message
    # Message text of error alert.
    # Available for type whole, decimal, date, time, textLength, list, custom
    # @see type
    # @return [String]
    # default nil
    attr_reader :error

    # Error Style (ST_DataValidationErrorStyle)
    # The style of error alert used for this data validation.
    # Options are:
    #  * information: This data validation error style uses an information icon in the error alert.
    #  * stop: This data validation error style uses a stop icon in the error alert.
    #  * warning: This data validation error style uses a warning icon in the error alert.
    # Available for type whole, decimal, date, time, textLength, list, custom
    # @see type
    # @return [Symbol]
    # default :stop
    attr_reader :errorStyle

    # Error Title
    # Title bar text of error alert.
    # Available for type whole, decimal, date, time, textLength, list, custom
    # @see type
    # @return [String]
    # default nil
    attr_reader :errorTitle

    # Operator (ST_DataValidationOperator)
    # The relational operator used with this data validation.
    # Options are:
    #  * between: Data validation which checks if a value is between two other values.
    #  * equal: Data validation which checks if a value is equal to a specified value.
    #  * greater_than: Data validation which checks if a value is greater than a specified value.
    #  * greater_than_or_equal: Data validation which checks if a value is greater than or equal to a specified value.
    #  * less_than: Data validation which checks if a value is less than a specified value.
    #  * less_than_or_equal: Data validation which checks if a value is less than or equal to a specified value.
    #  * not_between: Data validation which checks if a value is not between two other values.
    #  * not_equal: Data validation which checks if a value is not equal to a specified value.
    # Available for type whole, decimal, date, time, textLength
    # @see type
    # @return [Symbol]
    # default nil
    attr_reader :operator

    # Input prompt
    # Message text of input prompt.
    # Available for type whole, decimal, date, time, textLength, list, custom
    # @see type
    # @return [String]
    # default nil
    attr_reader :prompt

    # Prompt title
    # Title bar text of input prompt.
    # Available for type whole, decimal, date, time, textLength, list, custom
    # @see type
    # @return [String]
    # default nil
    attr_reader :promptTitle

    # Show drop down
    # A boolean value indicating whether to display a dropdown combo box for a list type data
    # validation. Be careful: false shows the dropdown list!
    # Available for type list
    # @see type
    # @return [Boolean]
    # default false
    attr_reader :showDropDown

    # Show error message
    # A boolean value indicating whether to display the error alert message when an invalid
    # value has been entered, according to the criteria specified.
    # Available for type whole, decimal, date, time, textLength, list, custom
    # @see type
    # @return [Boolean]
    # default false
    attr_reader :showErrorMessage

    # Show input message
    # A boolean value indicating whether to display the input prompt message.
    # Available for type whole, decimal, date, time, textLength, list, custom
    # @see type
    # @return [Boolean]
    # default false
    attr_reader :showInputMessage

    # Range over which data validation is applied, in "A1:B2" format
    # Available for type whole, decimal, date, time, textLength, list, custom
    # @see type
    # @return [String]
    # default nil
    attr_reader :sqref

    # The type (ST_DataValidationType) of data validation.
    # Options are:
    #  * custom: Data validation which uses a custom formula to check the cell value.
    #  * date: Data validation which checks for date values satisfying the given condition.
    #  * decimal: Data validation which checks for decimal values satisfying the given condition.
    #  * list: Data validation which checks for a value matching one of list of values.
    #  * none: No data validation.
    #  * textLength: Data validation which checks for text values, whose length satisfies the given condition.
    #  * time: Data validation which checks for time values satisfying the given condition.
    #  * whole: Data validation which checks for whole number values satisfying the given condition.
    # @return [Symbol]
    # default none
    attr_reader :type


    # @see formula1
    def formula1=(v); Axlsx::validate_string(v); @formula1 = v end

    # @see formula2
    def formula2=(v); Axlsx::validate_string(v); @formula2 = v end 

    # @see allowBlank
    def allowBlank=(v); Axlsx::validate_boolean(v); @allowBlank = v end

    # @see error
    def error=(v); Axlsx::validate_string(v); @error = v end

    # @see errorStyle
    def errorStyle=(v); Axlsx::validate_data_validation_error_style(v); @errorStyle = v end

    # @see errorTitle
    def errorTitle=(v); Axlsx::validate_string(v); @errorTitle = v end

    # @see operator
    def operator=(v); Axlsx::validate_data_validation_operator(v); @operator = v end

    # @see prompt
    def prompt=(v); Axlsx::validate_string(v); @prompt = v end

    # @see promptTitle
    def promptTitle=(v); Axlsx::validate_string(v); @promptTitle = v end

    # @see showDropDown
    def showDropDown=(v); Axlsx::validate_boolean(v); @showDropDown = v end

    # @see showErrorMessage
    def showErrorMessage=(v); Axlsx::validate_boolean(v); @showErrorMessage = v end

    # @see showInputMessage
    def showInputMessage=(v); Axlsx::validate_boolean(v); @showInputMessage = v end

    # @see sqref
    def sqref=(v); Axlsx::validate_string(v); @sqref = v end

    # @see type
    def type=(v); Axlsx::validate_data_validation_type(v); @type = v end

    # Serializes the data validation
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      valid_attributes = get_valid_attributes

      str << '<dataValidation '
      str << instance_values.map { |key, value| '' << key << '="' << value.to_s << '"' if (valid_attributes.include?(key.to_sym) and not CHILD_ELEMENTS.include?(key.to_sym)) }.join(' ')
      str << '>'
      str << '<formula1>' << self.formula1 << '</formula1>' if @formula1 and valid_attributes.include?(:formula1)
      str << '<formula2>' << self.formula2 << '</formula2>' if @formula2 and valid_attributes.include?(:formula2)
      str << '</dataValidation>'
    end

  private
    def get_valid_attributes
      attributes = [:allowBlank, :error, :errorStyle, :errorTitle, :prompt, :promptTitle, :showErrorMessage, :showInputMessage, :sqref, :type ]

      if [:whole, :decimal, :data, :time, :textLength].include?(@type)
        attributes << [:operator, :formula1]
        attributes << [:formula2] if [:between, :notBetween].include?(@operator)
      elsif @type == :list
        attributes << [:showDropDown, :formula1]
      elsif @type == :custom
        attributes << [:formula1]
      else
        attributes = []
      end

      attributes.flatten!
    end
  end
end
