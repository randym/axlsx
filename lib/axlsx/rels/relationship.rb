# encoding: UTF-8
module Axlsx
  # A relationship defines a reference between package parts.
  # @note Packages automatically manage relationships.
  class Relationship

    # The location of the relationship target
    # @return [String]
    attr_reader :Target

    # The type of relationship
    # @note Supported types are defined as constants in Axlsx:
    # @see XML_NS_R
    # @see TABLE_R
    # @see PIVOT_TABLE_R
    # @see WORKBOOK_R
    # @see WORKSHEET_R
    # @see APP_R
    # @see RELS_R
    # @see CORE_R
    # @see STYLES_R
    # @see CHART_R
    # @see DRAWING_R
    # @return [String]
    attr_reader :Type

    # The target mode of the relationship
    # used for hyperlink type relationships to mark the relationship to an external resource
    # TargetMode can be specified during initialization by passing in a :target_mode option
    # Target mode must be :external for now.
    attr_reader :TargetMode

    # creates a new relationship
    # @param [String] type The type of the relationship
    # @param [String] target The target for the relationship
    # @option [Symbol] :target_mode only accepts :external.
    def initialize(type, target, options={})
      self.Target=target
      self.Type=type
      self.TargetMode = options.delete(:target_mode) if options[:target_mode]
    end

    # @see Target
    def Target=(v) Axlsx::validate_string v; @Target = v end
    # @see Type
    def Type=(v) Axlsx::validate_relationship_type v; @Type = v end

    # @see TargetMode
    def TargetMode=(v) RestrictionValidator.validate 'Relationship.TargetMode', [:External, :Internal], v; @TargetMode = v; end

    # serialize relationship
    # @param [String] str
    # @param [Integer] rId the id for this relationship
    # @return [String]
    def to_xml_string(rId, str = '')
      h = self.instance_values
      h[:Id] = 'rId' << rId.to_s
      str << '<Relationship '
      str << h.map { |key, value| '' << key.to_s << '="' << Axlsx::coder.encode(value.to_s) << '"'}.join(' ')
      str << '/>'
    end

  end
end
