module Axlsx

  # A simple, self serializing class for storing pivot tables
  class PivotTables < SimpleTypedList

    # creates a new Tables object
    def initialize(worksheet)
      raise ArgumentError, "you must provide a worksheet" unless worksheet.is_a?(Worksheet)
      super PivotTable
      @worksheet = worksheet
    end

    # The worksheet that owns this collection of pivot tables
    # @return [Worksheet]
    attr_reader :worksheet

    # returns the relationships required by this collection
    def relationships
      return [] if empty?
      map{ |pivot_table| Relationship.new(pivot_table, PIVOT_TABLE_R, "../#{pivot_table.pn}") }
    end
  end

end
