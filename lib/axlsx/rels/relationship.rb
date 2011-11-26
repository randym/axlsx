module Axlsx
  # A relationship defines a reference between package parts.
  # @note Packages automatcially manage relationships.
  class Relationship

    # The location of the relationship target
    # @return [String]
    attr_reader :Target

    # The type of relationship
    # @note Supported types are defined as constants in Axlsx:
    # @see XML_NS_R
    # @see TABLE_R
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
    def initialize(type, target)
      self.Target=target
      self.Type=type
    end

    # @see Target
    def Target=(v) Axlsx::validate_string v; @Target = v end
    # @see Type
    def Type=(v) Axlsx::validate_relationship_type v; @Type = v end

    # Serializes the relationship    
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @param [String] rId the reference id of the object.
    # @return [String]
    def to_xml(xml, rId)
      h = self.instance_values
      h[:Id] = rId
      xml.Relationship(h)
    end
  end
end
