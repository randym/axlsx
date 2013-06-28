# encoding: UTF-8
module Axlsx
  # Validate a value against a specific list of allowed values.
  class RestrictionValidator
    # Perform validation
    # @param [String] name The name of what is being validatied. This is included in the error message
    # @param [Array] choices The list of choices to validate against
    # @param [Any] v The value to be validated
    # @raise [ArgumentError] Raised if the value provided is not in the list of choices.
    # @return [Boolean] true if validation succeeds.
    def self.validate(name, choices, v)
      raise ArgumentError, (ERR_RESTRICTION % [v.to_s, name, choices.inspect]) unless choices.include?(v)
      true
    end
  end

  # Validate that the value provided is between a specific range
  # Note that no data conversions will be done for you!
  # Comparisons will be made using < and > or <= and <= when the inclusive parameter is true
  class RangeValidator
    # @param [String] name The name of what is being validated
    # @param [Any] min The minimum allowed value
    # @param [Any] max The maximum allowed value
    # @param [Any] value The value to be validated
    # @param [Boolean] inclusive Flag indicating if the comparison should be inclusive.
    def self.validate(name, min, max, value, inclusive = true)
      passes = if inclusive
                 min <= value && value <= max
               else
                 min < value && value < max
               end
      raise ArgumentError, (ERR_RANGE % [value.inspect, min.to_s, max.to_s, inclusive]) unless passes
    end
  end
  # Validates the value against the regular expression provided.
  class RegexValidator
    # @param [String] name The name of what is being validated. This is included in the output when the value is invalid
    # @param [Regexp] regex The regular expression to evaluate
    # @param [Any] v The value to validate.
    def self.validate(name, regex, v)
      raise ArgumentError, (ERR_REGEX % [v.inspect, regex.to_s]) unless (v.respond_to?(:to_s) && v.to_s.match(regex))
    end
  end

  # Validate that the class of the value provided is either an instance or the class of the allowed types and that any specified additional validation returns true.
  class DataTypeValidator
    # Perform validation
    # @param [String] name The name of what is being validated. This is included in the error message
    # @param [Array, Class] types A single class or array of classes that the value is validated against.
    # @param [Block] other Any block that must evaluate to true for the value to be valid
    # @raise [ArugumentError] Raised if the class of the value provided is not in the specified array of types or the block passed returns false
    # @return [Boolean] true if validation succeeds.
    # @see validate_boolean
    def self.validate(name, types, v, other=false)
      types = [types] unless types.is_a? Array
      if other.is_a?(Proc)
         raise ArgumentError, (ERR_TYPE % [v.inspect, name, types.inspect]) unless other.call(v)
      end
      if v.class == Class
        types.each { |t| return if v.ancestors.include?(t) }
      else
        types.each { |t| return if v.is_a?(t) }
      end
      raise ArgumentError, (ERR_TYPE % [v.inspect, name, types.inspect])
    end
  end


  # Requires that the value can be converted to an integer
  # @para, [Any] v the value to validate
  # @raise [ArgumentError] raised if the value cannot be converted to an integer
  def self.validate_integerish(v)
    raise ArgumentError, (ERR_INTEGERISH % v.inspect) unless (v.respond_to?(:to_i) && v.to_i.is_a?(Integer))
  end

  # Requires that the value is between -54000000 and 54000000
  # @param [Any] v The value validated
  # @raise [ArgumentError] raised if the value cannot be converted to an integer between the allowed angle values for chart label rotation.
  # @return [Boolean] true if the data is valid
  def self.validate_angle(v)
    raise ArgumentError, (ERR_ANGLE % v.inspect) unless (v.to_i >= -5400000 && v.to_i <= 5400000)
  end
  # Requires that the value is a Fixnum or Integer and is greater or equal to 0
  # @param [Any] v The value validated
  # @raise [ArgumentError] raised if the value is not a Fixnum or Integer value greater or equal to 0
  # @return [Boolean] true if the data is valid
  def self.validate_unsigned_int(v)
    DataTypeValidator.validate(:unsigned_int, [Fixnum, Integer], v, lambda { |arg| arg.respond_to?(:>=) && arg >= 0 })
  end

  # Requires that the value is a Fixnum Integer or Float and is greater or equal to 0
  # @param [Any] v The value validated
  # @raise [ArgumentError] raised if the value is not a Fixnun, Integer, Float value greater or equal to 0
  # @return [Boolean] true if the data is valid
  def self.validate_unsigned_numeric(v)
    DataTypeValidator.validate("Invalid column width", [Fixnum, Integer, Float], v, lambda { |arg| arg.respond_to?(:>=) && arg.to_i >= 0 })
  end

  # Requires that the value is a Fixnum or Integer
  # @param [Any] v The value validated
  def self.validate_int(v)
    DataTypeValidator.validate :unsigned_int, [Fixnum, Integer], v
  end

  # Requires that the value is a form that can be evaluated as a boolean in an xml document.
  # The value must be an instance of Fixnum, String, Integer, Symbol, TrueClass or FalseClass and
  # it must be one of 0, 1, "true", "false", :true, :false, true, false, "0", or "1"
  # @param [Any] v The value validated
  def self.validate_boolean(v)
    DataTypeValidator.validate(:boolean, [Fixnum, String, Integer, Symbol, TrueClass, FalseClass], v, lambda { |arg| [0, 1, "true", "false", :true, :false, true, false, "0", "1"].include?(arg) })
  end

  # Requires that the value is a String
  # @param [Any] v The value validated
  def self.validate_string(v)
    DataTypeValidator.validate :string, String, v
  end

  # Requires that the value is a Float
  # @param [Any] v The value validated
  def self.validate_float(v)
    DataTypeValidator.validate :float, Float, v
  end

  # Requires that the value is a string containing a positive decimal number followed by one of the following units:
  # "mm", "cm", "in", "pt", "pc", "pi"
  def self.validate_number_with_unit(v)
    RegexValidator.validate "number_with_unit", /\A[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)\Z/, v
  end

  # Requires that the value is an integer ranging from 10 to 400.
  def self.validate_scale_10_400(v)
    DataTypeValidator.validate "page_scale", [Fixnum, Integer], v, lambda { |arg| arg >= 10 && arg <= 400 }
  end

  # Requires that the value is an integer ranging from 10 to 400 or 0.
  def self.validate_scale_0_10_400(v)
    DataTypeValidator.validate "page_scale", [Fixnum, Integer], v, lambda { |arg| arg == 0 || (arg >= 10 && arg <= 400) }
  end

  # Requires that the value is one of :default, :landscape, or :portrait.
  def self.validate_page_orientation(v)
    RestrictionValidator.validate "page_orientation", [:default, :landscape, :portrait], v
  end
  # Requires that the value is one of :none, :single, :double, :singleAccounting, :doubleAccounting
  def self.validate_cell_u(v)
    RestrictionValidator.validate "cell run style u", [:none, :single, :double, :singleAccounting, :doubleAccounting], v
  end

  # validates cell style family which must be between 1 and 5 
  def self.validate_family(v)
    RestrictionValidator.validate "cell run style family", 1..5, v
  end
  # Requires that the value is valid pattern type.
  # valid pattern types must be one of :none, :solid, :mediumGray, :darkGray, :lightGray, :darkHorizontal, :darkVertical, :darkDown,
  # :darkUp, :darkGrid, :darkTrellis, :lightHorizontal, :lightVertical, :lightDown, :lightUp, :lightGrid, :lightTrellis, :gray125, or :gray0625.
  # @param [Any] v The value validated
  def self.validate_pattern_type(v)
    RestrictionValidator.validate :pattern_type, [:none, :solid, :mediumGray, :darkGray, :lightGray, :darkHorizontal, :darkVertical, :darkDown, :darkUp, :darkGrid,
                                                  :darkTrellis, :lightHorizontal, :lightVertical, :lightDown, :lightUp, :lightGrid, :lightTrellis, :gray125, :gray0625], v
  end

  # Requires that the value is one of the ST_TimePeriod types
  # valid time period types are today, yesterday, tomorrow, last7Days,
  # thisMonth, lastMonth, nextMonth, thisWeek, lastWeek, nextWeek
  def self.validate_time_period_type(v)
    RestrictionValidator.validate :time_period_type, [:today, :yesterday, :tomorrow, :last7Days, :thisMonth, :lastMonth, :nextMonth, :thisWeek, :lastWeek, :nextWeek], v
  end

  # Requires that the value is one of the valid ST_IconSet types
  # Allowed values are: 3Arrows, 3ArrowsGray, 3Flags, 3TrafficLights1, 3TrafficLights2, 3Signs, 3Symbols, 3Symbols2, 4Arrows, 4ArrowsGray, 4RedToBlack, 4Rating, 4TrafficLights, 5Arrows, 5ArrowsGray, 5Rating, 5Quarters
  def self.validate_icon_set(v)
    RestrictionValidator.validate :iconSet, ["3Arrows", "3ArrowsGray", "3Flags", "3TrafficLights1", "3TrafficLights2", "3Signs", "3Symbols", "3Symbols2", "4Arrows", "4ArrowsGray", "4RedToBlack", "4Rating", "4TrafficLights", "5Arrows", "5ArrowsGray", "5Rating", "5Quarters"], v
  end

  # Requires that the value is valid conditional formatting type.
  # valid types must be one of expression, cellIs, colorScale,
  # dataBar, iconSet, top10, uniqueValues, duplicateValues,
  # containsText, notContainsText, beginsWith, endsWith,
  # containsBlanks, notContainsBlanks, containsErrors,
  # notContainsErrors, timePeriod, aboveAverage
  # @param [Any] v The value validated
  def self.validate_conditional_formatting_type(v)
    RestrictionValidator.validate :conditional_formatting_type, [:expression, :cellIs, :colorScale, :dataBar, :iconSet, :top10, :uniqueValues, :duplicateValues, :containsText, :notContainsText, :beginsWith, :endsWith, :containsBlanks, :notContainsBlanks, :containsErrors, :notContainsErrors, :timePeriod, :aboveAverage], v
  end

  # Requires thatt he value is a valid conditional formatting value object type.
  # valid types must be one of num, percent, max, min, formula and percentile
  def self.validate_conditional_formatting_value_object_type(v)
    RestrictionValidator.validate :conditional_formatting_value_object_type, [:num, :percent, :max, :min, :formula, :percentile], v
  end

  # Requires that the value is valid conditional formatting operator.
  # valid operators must be one of lessThan, lessThanOrEqual, equal,
  # notEqual, greaterThanOrEqual, greaterThan, between, notBetween,
  # containsText, notContains, beginsWith, endsWith
  # @param [Any] v The value validated
  def self.validate_conditional_formatting_operator(v)
    RestrictionValidator.validate :conditional_formatting_type, [:lessThan, :lessThanOrEqual, :equal, :notEqual, :greaterThanOrEqual, :greaterThan, :between, :notBetween, :containsText, :notContains, :beginsWith, :endsWith], v
  end

  # Requires that the value is a gradient_type.
  # valid types are :linear and :path
  # @param [Any] v The value validated
  def self.validate_gradient_type(v)
    RestrictionValidator.validate :gradient_type, [:linear, :path], v
  end

  # Requires that the value is a valid scatterStyle
  # must be one of :none | :line | :lineMarker | :marker | :smooth | :smoothMarker
  # must be one of "none" | "line" | "lineMarker" | "marker" | "smooth" | "smoothMarker"
  # @param [Symbol|String] v the value to validate
  def self.validate_scatter_style(v)
    Axlsx::RestrictionValidator.validate "ScatterChart.scatterStyle", [:none, :line, :lineMarker, :marker, :smooth, :smoothMarker], v.to_sym
  end
  # Requires that the value is a valid horizontal_alignment
  # :general, :left, :center, :right, :fill, :justify, :centerContinuous, :distributed are allowed
  # @param [Any] v The value validated
  def self.validate_horizontal_alignment(v)
    RestrictionValidator.validate :horizontal_alignment, [:general, :left, :center, :right, :fill, :justify, :centerContinuous, :distributed], v
  end

  # Requires that the value is a valid vertical_alignment
  # :top, :center, :bottom, :justify, :distributed are allowed
  # @param [Any] v The value validated
  def self.validate_vertical_alignment(v)
    RestrictionValidator.validate :vertical_alignment, [:top, :center, :bottom, :justify, :distributed], v
  end

  # Requires that the value is a valid content_type
  # TABLE_CT, WORKBOOK_CT, APP_CT, RELS_CT, STYLES_CT, XML_CT, WORKSHEET_CT, SHARED_STRINGS_CT, CORE_CT, CHART_CT, DRAWING_CT, COMMENT_CT are allowed
  # @param [Any] v The value validated
  def self.validate_content_type(v)
    RestrictionValidator.validate :content_type, [TABLE_CT, WORKBOOK_CT, APP_CT, RELS_CT, STYLES_CT, XML_CT, WORKSHEET_CT, SHARED_STRINGS_CT, CORE_CT, CHART_CT, JPEG_CT, GIF_CT, PNG_CT, DRAWING_CT, COMMENT_CT, VML_DRAWING_CT, PIVOT_TABLE_CT, PIVOT_TABLE_CACHE_DEFINITION_CT], v
  end

  # Requires that the value is a valid relationship_type
  # XML_NS_R, TABLE_R, WORKBOOK_R, WORKSHEET_R, APP_R, RELS_R, CORE_R, STYLES_R, CHART_R, DRAWING_R, IMAGE_R, HYPERLINK_R, SHARED_STRINGS_R are allowed
  # @param [Any] v The value validated
  def self.validate_relationship_type(v)
    RestrictionValidator.validate :relationship_type, [XML_NS_R, TABLE_R, WORKBOOK_R, WORKSHEET_R, APP_R, RELS_R, CORE_R, STYLES_R, CHART_R, DRAWING_R, IMAGE_R, HYPERLINK_R, SHARED_STRINGS_R, COMMENT_R, VML_DRAWING_R, COMMENT_R_NULL, PIVOT_TABLE_R, PIVOT_TABLE_CACHE_DEFINITION_R], v
  end

  # Requires that the value is a valid table element type
  # :wholeTable, :headerRow, :totalRow, :firstColumn, :lastColumn, :firstRowStripe, :secondRowStripe, :firstColumnStripe, :secondColumnStripe, :firstHeaderCell, :lastHeaderCell, :firstTotalCell, :lastTotalCell, :firstSubtotalColumn, :secondSubtotalColumn, :thirdSubtotalColumn, :firstSubtotalRow, :secondSubtotalRow, :thirdSubtotalRow, :blankRow, :firstColumnSubheading, :secondColumnSubheading, :thirdColumnSubheading, :firstRowSubheading, :secondRowSubheading, :thirdRowSubheading, :pageFieldLabels, :pageFieldValues are allowed
  # @param [Any] v The value validated
  def self.validate_table_element_type(v)
    RestrictionValidator.validate :table_element_type, [:wholeTable, :headerRow, :totalRow, :firstColumn, :lastColumn, :firstRowStripe, :secondRowStripe, :firstColumnStripe, :secondColumnStripe, :firstHeaderCell, :lastHeaderCell, :firstTotalCell, :lastTotalCell, :firstSubtotalColumn, :secondSubtotalColumn, :thirdSubtotalColumn, :firstSubtotalRow, :secondSubtotalRow, :thirdSubtotalRow, :blankRow, :firstColumnSubheading, :secondColumnSubheading, :thirdColumnSubheading, :firstRowSubheading, :secondRowSubheading, :thirdRowSubheading, :pageFieldLabels, :pageFieldValues], v
  end

  # Requires that the value is a valid data_validation_error_style
  # :information, :stop, :warning
  # @param [Any] v The value validated
  def self.validate_data_validation_error_style(v)
    RestrictionValidator.validate :validate_data_validation_error_style, [:information, :stop, :warning], v
  end

  # Requires that the value is valid data validation operator.
  # valid operators must be one of lessThan, lessThanOrEqual, equal,
  # notEqual, greaterThanOrEqual, greaterThan, between, notBetween
  # @param [Any] v The value validated
  def self.validate_data_validation_operator(v)
    RestrictionValidator.validate :data_validation_operator, [:lessThan, :lessThanOrEqual, :equal, :notEqual, :greaterThanOrEqual, :greaterThan, :between, :notBetween], v
  end

  # Requires that the value is valid data validation type.
  # valid types must be one of custom, data, decimal, list, none, textLength, time, whole
  # @param [Any] v The value validated
  def self.validate_data_validation_type(v)
    RestrictionValidator.validate :data_validation_type, [:custom, :data, :decimal, :list, :none, :textLength, :time, :whole], v
  end

  # Requires that the value is a valid sheet view type.
  # valid types must be one of normal, page_break_preview, page_layout
  # @param [Any] v The value validated
  def self.validate_sheet_view_type(v)
    RestrictionValidator.validate :sheet_view_type, [:normal, :page_break_preview, :page_layout], v
  end

  # Requires that the value is a valid active pane type.
  # valid types must be one of bottom_left, bottom_right, top_left, top_right
  # @param [Any] v The value validated
  def self.validate_pane_type(v)
    RestrictionValidator.validate :active_pane_type, [:bottom_left, :bottom_right, :top_left, :top_right], v
  end

  # Requires that the value is a valid split state type.
  # valid types must be one of frozen, frozen_split, split
  # @param [Any] v The value validated
  def self.validate_split_state_type(v)
    RestrictionValidator.validate :split_state_type, [:frozen, :frozen_split, :split], v
  end

  # Requires that the value is a valid "display blanks as" type.
  # valid types must be one of gap, span, zero
  # @param [Any] v The value validated
  def self.validate_display_blanks_as(v)
    RestrictionValidator.validate :display_blanks_as, [:gap, :span, :zero], v
  end
end
