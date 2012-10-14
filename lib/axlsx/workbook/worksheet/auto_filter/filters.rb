module Axlsx

  # When multiple values are chosen to filter by, or when a group of date values are chosen to filter by, 
  # this object groups those criteria together.
  class Filters
    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes

    # Creates a new Filters object
    # @param [Hash] options Options used to set this objects attributes and
    #                       create filter and/or date group items
    # @option [Boolean] blank @see blank
    # @option [String] calendar_type @see calendar_type
    # @option [Array] filter_items An array of values that will be used to create filter objects.
    # @option [Array] date_group_items An array of hases defining date group item filters to apply.
    # @note The recommended way to interact with filter objects is via AutoFilter#add_column
    # @example
    #   ws.auto_filter.add_column(0, :filters, :blank => true, :calendar_type => 'japan', :filter_items => [100, 'a'])
    def initialize(options={})
      parse_options options
    end

    serializable_attributes :blank, :calendar_type

    # Allowed calendar types
    CALENDAR_TYPES = %w(gregorian gregorianUs gregorianMeFrench gregorianArabic hijri hebrew taiwan japan thai korea saka gregorianXlitEnglish gregorianXlitFrench none)

    # Flag indicating whether to filter by blank.
    # @return [Boolean]
    attr_reader :blank

    # Calendar type for date grouped items. 
    # Used to interpret the values in dateGroupItem.
    # This is the calendar type used to evaluate all dates in the filter column,
    # even when those dates are not using the same calendar system / date formatting.
    attr_reader :calendar_type

    # Tells us if the row of the cell provided should be filterd as it
    # does not meet any of the specified filter_items or
    # date_group_items restrictions.
    # @param [Cell] cell The cell to test against items
    # TODO implement this for date filters as well!
    def apply(cell)
      return false unless cell
      filter_items.each do |filter|
        return false if cell.value == filter.val
      end
      true
    end

    # The filter values in this filters object
    def filter_items
      @filter_items ||= []
    end

    # the date group values in this filters object
    def date_group_items
      @date_group_items ||= []
    end

    # @see calendar_type
    # @param [String] calendar The calendar type to use. This must be one of the types defined in CALENDAR_TYPES
    # @return [String]
    def calendar_type=(calendar)
      RestrictionValidator.validate 'Filters.calendar_type', CALENDAR_TYPES, calendar
      @calendar_type = calendar
    end

    # Set the value for blank
    # @see blank
    def blank=(use_blank)
      Axlsx.validate_boolean use_blank
      @blank = use_blank
    end

    # Serialize the object to xml
    def to_xml_string(str = '')
      str << "<filters #{serialized_attributes}>"
      filter_items.each {  |filter| filter.to_xml_string(str) }
      date_group_items.each { |date_group_item| date_group_item.to_xml_string(str) }
      str << '</filters>'
    end

    # not entirely happy with this. 
    # filter_items should be a simple typed list that overrides << etc
    # to create Filter objects from the inserted values. However this
    # is most likely so rarely used...(really? do you know that?)
    def filter_items=(values)
      values.each do |value|
        filter_items << Filter.new(value)
      end
    end

    # Date group items are date group filter items where you specify the
    # date_group and a value for that option as part of the auto_filter
    # @note This can be specified, but will not be applied to the date
    # values in your workbook at this time.
    def date_group_items=(options)
      options.each do |date_group|
        raise ArgumentError, "date_group_items should be an array of hashes specifying the options for each date_group_item" unless date_group.is_a?(Hash)
        date_group_items << DateGroupItem.new(date_group)
      end
    end

    # This class expresses a filter criteria value.
    class Filter

      # Creates a new filter value object
      # @param [Any] value   The value of the filter. This is not restricted, but
      #                       will be serialized via to_s so if you are passing an object
      #                       be careful.
      def initialize(value)
        @val = value
      end


      #Filter value used in the criteria.
      attr_accessor :val

      # Serializes the filter value object
      # @param [String] str The string to concact the serialization information to.
      def to_xml_string(str = '')
        str << "<filter val='#{@val.to_s}' />"
      end
    end


    # This collection is used to express a group of dates or times which are
    # used in an AutoFilter criteria. Values are always written in the calendar
    # type of the first date encountered in the filter range, so that all
    # subsequent dates, even when formatted or represented by other calendar
    # types, can be correctly compared for the purposes of filtering.
    class DateGroupItem
      include Axlsx::OptionsParser
include Axlsx::SerializedAttributes

      # Creates a new DateGroupItem
      # @param [Hash] options A hash of options to use when
      # instanciating the object
      # @option [String] date_time_grouping the part of the date this
      # filter should apply for grouping
      # @option [Integer|String] year @see year
      # @option [Integer] month @see month
      # @option [Integer] day @see day
      # @option [Integer] hour @see hour
      # @option [Integer] minute @see minute
      # @option [Integer] second @see second
      def initialize(options={})
        raise ArgumentError,  "You must specify a year for date time grouping" unless options[:year]
        raise ArgumentError, "You must specify a date_time_grouping when creating a DateGroupItem for auto filter" unless options[:date_time_grouping]
        parse_options options
      end

      serializable_attributes :date_time_grouping, :year, :month, :day, :hour, :minute, :second

      # Allowed date time groupings
      DATE_TIME_GROUPING = %w(year month day hour minute second)

      # Grouping level
      # This must be one of year, month, day, hour, minute or second.
      # @return [String]
      attr_reader :date_time_grouping

      # Year (4 digits)
      # @return [Integer|String]
      attr_reader :year

      # Month (1..12)
      # @return [Integer]
      attr_reader :month

      # Day (1-31)
      # @return [Integer]
      attr_reader :day

      # Hour (0..23)
      # @return [Integer]
      attr_reader :hour

      # Minute (0..59(
      # @return [Integer]
      attr_reader :minute

      # Second (0..59)
      # @return [Integer]
      attr_reader :second

      # The year value for the date group item
      # This must be a four digit value
      def year=(value)
        RegexValidator.validate "DateGroupItem.year", /\d{4}/, value
        @year = value
      end

      # The month value for the date group item
      # This must be between 1 and 12
      def month=(value)
        RangeValidator.validate "DateGroupItem.month", 0, 12, value
        @month = value
      end

      # The day value for the date group item
      # This must be between 1 and 31 
      # @note no attempt is made to ensure the date value is valid for any given month
      def day=(value)
        RangeValidator.validate "DateGroupItem.day", 0, 31, value
        @day = value
      end

      # The hour value for the date group item
      # # this must be between 0 and 23
      def hour=(value)
        RangeValidator.validate "DateGroupItem.hour", 0, 23, value
        @hour = value
      end

      # The minute value for the date group item
      # This must be between 0 and 59
      def minute=(value)
        RangeValidator.validate "DateGroupItem.minute", 0, 59, value
        @minute = value
      end

      # The second value for the date group item
      # This must be between 0 and 59
      def second=(value)
        RangeValidator.validate "DateGroupItem.second", 0, 59, value
        @second = value
      end

      # The date time grouping for this filter.
      def date_time_grouping=(grouping)
        RestrictionValidator.validate 'DateGroupItem.date_time_grouping', DATE_TIME_GROUPING, grouping.to_s
        @date_time_grouping = grouping.to_s
      end

      # Serialize the object to xml
      # @param [String] str The string object this serialization will be concatenated to.
      def to_xml_string(str = '')
        str << '<dateGroupItem '
        serialized_attributes str
        str << '/>'
      end
    end
  end
end
