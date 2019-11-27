# frozen_string_literal: true
module Axlsx

  #A collection of Cfvo objects that initializes with the required
  #first two items
  class Cfvos < SimpleTypedList

    def initialize
      super(Cfvo)
    end

    # Serialize the Cfvo object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = String.new)
      each { |cfvo| cfvo.to_xml_string(str) }
    end
  end
end
