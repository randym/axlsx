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

    # The relationship instance for the given source object, or nil if none exists.
    # @see Relationship#source_obj
    # @return [Relationship]
    def for(source_obj)
      find{ |rel| rel.source_obj == source_obj }
    end

    # serialize relationships
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<?xml version="1.0" encoding="UTF-8"?>'
      str << ('<Relationships xmlns="' << RELS_R << '">')
      each{ |rel| rel.to_xml_string(str) }
      str << '</Relationships>'
    end

  end
end
