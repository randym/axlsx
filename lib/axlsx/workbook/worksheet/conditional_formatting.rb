module Axlsx
  # Conditional formatting allows styling of ranges based on functions
  #
  # @note The recommended way to manage conditional formatting is via Worksheet#add_conditional_formatting
  # @see Worksheet#add_conditional_formatting
  class ConditionalFormatting

    # Sqref
    # Range over which the formatting is applied
    # @return [String]
    attr_reader :sqref

    # Rules
    # Rules to apply the formatting to. Can be either a hash of
    # options for one ConditionalFormattingRule, an array of hashes
    # for multiple ConditionalFormattingRules, or an array of already
    # created ConditionalFormattingRules.
    # @return [Array]
    attr_reader :rules
    
    # Creates a new ConditionalFormatting object
    # @option options [Array] rules The rules to apply
    # @option options [String] sqref The range to apply the rules to
    def initialize(options={})
      @rules = []
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
    end

    # Add ConditionalFormattingRules to this object. Rules can either
    # be already created objects or hashes of options for automatic
    # creation. 
    # @option rules [Array, Hash] the rules apply, can be just one in hash form
    # @see ConditionalFormattingRule#initialize    
    def add_rules(rules)
      rules = [rules] if rules.is_a? Hash
      conditional_rules = rules.each do |rule|
        add_rule rule
      end
    end

    # Add a ConditionalFormattingRule. If a hash of options is passed
    # in create a rule on the fly
    # @option rule [ConditionalFormattingRule, Hash] A rule to create
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
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<conditionalFormatting sqref="' << sqref << '">'
      str << rules.collect{ |rule| rule.to_xml_string }.join(' ')
      str << '</conditionalFormatting>'
    end
  end
end
