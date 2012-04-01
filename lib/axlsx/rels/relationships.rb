# encoding: UTF-8
module Axlsx
require 'axlsx/rels/relationship.rb'

  # Relationships are a collection of Relations that define how package parts are related.
  # @note The package automatically manages releationships.
  class Relationships < SimpleTypedList

    # Creates a new Relationships collection based on SimpleTypedList
    def initialize
      super Relationship
    end

    def to_xml_string(str = '')
      str << '<?xml version="1.0" encoding="UTF-8"?>'
      str << '<Relationships xmlns="' << RELS_R << '">'
      each_with_index { |rel, index| rel.to_xml_string(index+1, str) }
      str << '</Relationships>'
    end
    # Serializes the relationships document.
    # @return [String]
    def to_xml()
      builder = Nokogiri::XML::Builder.new(:encoding => ENCODING) do |xml|
        xml.Relationships(:xmlns => Axlsx::RELS_R) {
          each_with_index { |rel, index| rel.to_xml(xml, "rId#{index+1}") }
        }
      end
      builder.to_xml(:save_with => 0)
    end

  end
end
