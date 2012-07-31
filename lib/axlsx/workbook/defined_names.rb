module Axlsx

  class DefinedNames < SimpleTypedList

    def initialize
      super DefinedName
    end

    def to_xml_string(str = '')
      return if @list.empty?
      str << "<definedNames>"
      each { |defined_name| defined_name.to_xml_string(str) }
      str << '</definedNames>'
    end
  end
end

