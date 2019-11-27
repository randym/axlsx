# frozen_string_literal: true
module Axlsx
  # a simple types list of DefinedName objects
  class DefinedNames < SimpleTypedList

    # creates the DefinedNames object
    def initialize
      super DefinedName
    end

    # Serialize to xml
    # @param [String] str
    # @return [String]
    def to_xml_string(str = String.new)
      return if empty?
      str << '<definedNames>'
      each { |defined_name| defined_name.to_xml_string(str) }
      str << '</definedNames>'
    end
  end
end

