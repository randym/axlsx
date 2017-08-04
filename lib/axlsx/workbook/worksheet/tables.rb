module Axlsx

  # A simple, self serializing class for storing tables
  class Tables < SimpleTypedList

    # creates a new Tables object
    def initialize(worksheet)
      raise ArgumentError, "you must provide a worksheet" unless worksheet.is_a?(Worksheet)
      super Table
      @worksheet = worksheet
    end

    # The worksheet that owns this collection of tables
    # @return [Worksheet]
    attr_reader :worksheet

    # returns the relationships required by this collection
    def relationships
      return [] if empty?
      map{ |table| Relationship.new(table, TABLE_R, "../#{table.pn}") }
    end

    # renders the tables xml
    # @param [String] str
    # @return [String]
    def to_xml_string(str = "")
      return if empty?
      str << "<tableParts count='#{size}'>"
      each { |table| str << "<tablePart r:id='#{table.rId}'/>" }
      str << '</tableParts>'
    end
  end

end
