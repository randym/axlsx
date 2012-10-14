module Axlsx
  # Conditional formatting allows styling of ranges based on functions
  #
  # @note The recommended way to manage conditional formatting is via Worksheet#add_conditional_formatting
  # @see Worksheet#add_conditional_formatting
  # @see ConditionalFormattingRule
  class ConditionalFormatting

   include Axlsx::OptionsParser
   
    # Creates a new {ConditionalFormatting} object
    # @option options [Array] rules The rules to apply
    # @option options [String] sqref The range to apply the rules to
    def initialize(options={})
      @rules = []
      parse_options options
    end

    # Range over which the formatting is applied, in "A1:B2" format
    # @return [String]
    attr_reader :sqref

    # Rules to apply the formatting to. Can be either a hash of
    # options to create a {ConditionalFormattingRule}, an array of hashes
    # for multiple ConditionalFormattingRules, or an array of already
    # created ConditionalFormattingRules.
    # @see ConditionalFormattingRule#initialize
    # @return [Array]
    attr_reader :rules

     # Add Conditional Formatting Rules to this object. Rules can either
    # be already created {ConditionalFormattingRule} elements or
    # hashes of options for automatic creation.  If rules is a hash
    # instead of an array, assume only one rule being added.
    #
    # @example This would apply formatting "1" to cells > 20, and formatting "2" to cells < 1
    #        conditional_formatting.add_rules [
    #            { :type => :cellIs, :operator => :greaterThan, :formula => "20", :dxfId => 1, :priority=> 1 },
    #            { :type => :cellIs, :operator => :lessThan, :formula => "10", :dxfId => 2, :priority=> 2 } ]
    #
    # @param [Array|Hash] rules the rules to apply, can be just one in hash form
    # @see ConditionalFormattingRule#initialize
    def add_rules(rules)
      rules = [rules] if rules.is_a? Hash
      rules.each do |rule|
        add_rule rule
      end
    end

    # Add a ConditionalFormattingRule. If a hash of options is passed
    # in create a rule on the fly.
    # @param [ConditionalFormattingRule|Hash] rule A rule to use, or the options necessary to create one.
    # @see ConditionalFormattingRule#initialize
    def add_rule(rule)
      if rule.is_a? Axlsx::ConditionalFormattingRule
        @rules << rule
      elsif rule.is_a? Hash
        @rules << ConditionalFormattingRule.new(rule)
      end
    end

    # @see rules
    def rules=(v); @rules = v end
    # @see sqref
    def sqref=(v); Axlsx::validate_string(v); @sqref = v end

    # Serializes the conditional formatting element
    # @example Conditional Formatting XML looks like:
    #    <conditionalFormatting sqref="E3:E9">
    #        <cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan">
    #             <formula>0.5</formula>
    #        </cfRule>
    #    </conditionalFormatting>
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<conditionalFormatting sqref="' << sqref << '">'
      str << rules.collect{ |rule| rule.to_xml_string }.join(' ')
      str << '</conditionalFormatting>'
    end
  end
end
