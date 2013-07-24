# -*- coding: utf-8 -*-
module Axlsx
  # Conditional formatting rules specify formulas whose evaluations
  # format cells
  #
  # @note The recommended way to manage these rules is via Worksheet#add_conditional_formatting
  # @see Worksheet#add_conditional_formatting
  # @see ConditionalFormattingRule#initialize
  class ConditionalFormattingRule

    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes

    # Creates a new Conditional Formatting Rule object
    # @option options [Symbol] type The type of this formatting rule
    # @option options [Boolean] aboveAverage This is an aboveAverage rule
    # @option options [Boolean] bottom This is a bottom N rule.
    # @option options [Integer] dxfId The formatting id to apply to matches
    # @option options [Boolean] equalAverage Is the aboveAverage or belowAverage rule inclusive
    # @option options [Integer] priority The priority of the rule, 1 is highest
    # @option options [Symbol] operator Which operator to apply
    # @option options [String] text The value to apply a text operator against
    # @option options [Boolean] percent If a top/bottom N rule, evaluate as N% rather than N
    # @option options [Integer] rank If a top/bottom N rule, the value of N
    # @option options [Integer] stdDev The number of standard deviations above or below the average to match
    # @option options [Boolean] stopIfTrue Stop evaluating rules after this rule matches
    # @option options [Symbol]  timePeriod The time period in a date occuring... rule
    # @option options [String] formula The formula to match against in i.e. an equal rule. Use a [minimum, maximum] array for cellIs between/notBetween conditionals.
    def initialize(options={})
      @color_scale = @data_bar = @icon_set = @formula = nil
      parse_options options
    end

    serializable_attributes :type, :aboveAverage, :bottom, :dxfId, :equalAverage,
                            :priority, :operator, :text, :percent, :rank, :stdDev,
                            :stopIfTrue, :timePeriod

    # Formula
    # The formula or value to match against (e.g. 5 with an operator of :greaterThan to specify cell_value > 5).
    # If the operator is :between or :notBetween, use an array to specify [minimum, maximum]
    # @return [String]
    attr_reader :formula

    # Type (ST_CfType)
    # options are expression, cellIs, colorScale, dataBar, iconSet,
    # top10, uniqueValues, duplicateValues, containsText,
    # notContainsText, beginsWith, endsWith, containsBlanks,
    # notContainsBlanks, containsErrors, notContainsErrors,
    # timePeriod, aboveAverage
    # @return [Symbol]
    attr_reader :type

    # Above average rule
    # Indicates whether the rule is an "above average" rule. True
    # indicates 'above average'. This attribute is ignored if type is
    # not equal to aboveAverage.
    # @return [Boolean]
    attr_reader :aboveAverage

    # Bottom N rule
    # @return [Boolean]
    attr_reader :bottom

    # Differential Formatting Id
    # @return [Integer]
    attr_reader :dxfId

    # Equal Average
    # Flag indicating whether the 'aboveAverage' and 'belowAverage'
    # criteria is inclusive of the average itself, or exclusive of
    # that value.
    # @return [Boolean]
    attr_reader :equalAverage

    # Operator
    # The operator in a "cell value is" conditional formatting
    # rule. This attribute is ignored if type is not equal to cellIs
    #
    # Operator must be one of lessThan, lessThanOrEqual, equal,
    # notEqual, greaterThanOrEqual, greaterThan, between, notBetween,
    # containsText, notContains, beginsWith, endsWith
    # @return [Symbol]
    attr_reader :operator

    # Priority
    # The priority of this conditional formatting rule. This value is
    # used to determine which format should be evaluated and
    # rendered. Lower numeric values are higher priority than higher
    # numeric values, where '1' is the highest priority.
    # @return [Integer]
    attr_reader :priority

    # Text
    # used in a "text contains" conditional formatting
    # rule.
    # @return [String]
    attr_reader :text

    # percent (Top 10 Percent)
    # indicates whether a "top/bottom n" rule is a "top/bottom n
    # percent" rule. This attribute is ignored if type is not equal to
    # top10.
    # @return [Boolean]
    attr_reader :percent

    # rank (Rank)
    # The value of "n" in a "top/bottom n" conditional formatting
    # rule. This attribute is ignored if type is not equal to top10.
    # @return [Integer]
    attr_reader :rank

    # stdDev (StdDev)
    # The number of standard deviations to include above or below the
    # average in the conditional formatting rule. This attribute is
    # ignored if type is not equal to aboveAverage. If a value is
    # present for stdDev and the rule type = aboveAverage, then this
    # rule is automatically an "above or below N standard deviations"
    # rule.
    # @return [Integer]
    attr_reader :stdDev

    # stopIfTrue (Stop If True)
    # If this flag is '1', no rules with lower priority shall be
    # applied over this rule, when this rule evaluates to true.
    # @return [Boolean]
    attr_reader :stopIfTrue

    # timePeriod (Time Period)
    # The applicable time period in a "date occurring…" conditional
    # formatting rule. This attribute is ignored if type is not equal
    # to timePeriod.
    # Valid types are today, yesterday, tomorrow, last7Days,
    # thisMonth, lastMonth, nextMonth, thisWeek, lastWeek, nextWeek
    attr_reader :timePeriod


    # colorScale (Color Scale)
    # The color scale to apply to this conditional formatting
    # @return [ColorScale]
    def color_scale
      @color_scale ||= ColorScale.new
    end

    # dataBar (Data Bar)
    # The data bar to apply to this conditional formatting
    # @return [DataBar]
    def data_bar
      @data_bar ||= DataBar.new
    end

    # iconSet (Icon Set)
    # The icon set to apply to this conditional formatting
    # @return [IconSet]
    def icon_set
      @icon_set ||= IconSet.new
    end

    # @see type
    def type=(v); Axlsx::validate_conditional_formatting_type(v); @type = v end
    # @see aboveAverage
    def aboveAverage=(v); Axlsx::validate_boolean(v); @aboveAverage = v end
    # @see bottom
    def bottom=(v); Axlsx::validate_boolean(v); @bottom = v end
    # @see dxfId
    def dxfId=(v); Axlsx::validate_unsigned_numeric(v); @dxfId = v end
    # @see equalAverage
    def equalAverage=(v); Axlsx::validate_boolean(v); @equalAverage = v end
    # @see priority
    def priority=(v); Axlsx::validate_unsigned_numeric(v); @priority = v end
    # @see operator
    def operator=(v); Axlsx::validate_conditional_formatting_operator(v); @operator = v end
    # @see text
    def text=(v); Axlsx::validate_string(v); @text = v end
    # @see percent
    def percent=(v); Axlsx::validate_boolean(v); @percent = v end
    # @see rank
    def rank=(v); Axlsx::validate_unsigned_numeric(v); @rank = v end
    # @see stdDev
    def stdDev=(v); Axlsx::validate_unsigned_numeric(v); @stdDev = v end
    # @see stopIfTrue
    def stopIfTrue=(v); Axlsx::validate_boolean(v); @stopIfTrue = v end
    # @see timePeriod
    def timePeriod=(v); Axlsx::validate_time_period_type(v); @timePeriod = v end
    # @see formula
    def formula=(v); [*v].each {|x| Axlsx::validate_string(x) }; @formula = [*v].map { |form| ::CGI.escapeHTML(form) } end

    # @see color_scale
    def color_scale=(v)
      Axlsx::DataTypeValidator.validate 'conditional_formatting_rule.color_scale', ColorScale, v
      @color_scale = v
    end

    # @see data_bar
    def data_bar=(v)
      Axlsx::DataTypeValidator.validate 'conditional_formatting_rule.data_bar', DataBar, v
      @data_bar = v
    end

    # @see icon_set
    def icon_set=(v)
      Axlsx::DataTypeValidator.validate 'conditional_formatting_rule.icon_set', IconSet, v
      @icon_set = v
    end


    # Serializes the conditional formatting rule
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<cfRule '
      serialized_attributes str
      str << '>'
      str << '<formula>' << [*self.formula].join('</formula><formula>') << '</formula>' if @formula
      @color_scale.to_xml_string(str) if @color_scale && @type == :colorScale
      @data_bar.to_xml_string(str) if @data_bar && @type == :dataBar
      @icon_set.to_xml_string(str) if @icon_set && @type == :iconSet
      str << '</cfRule>'
    end
  end
end
