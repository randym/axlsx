module Axlsx

  # A simple, self serializing class for storing conditional formattings
  class ConditionalFormattings < SimpleTypedList

    # creates a new Tables object
    def initialize(worksheet)
      raise ArgumentError, "you must provide a worksheet" unless worksheet.is_a?(Worksheet)
      super ConditionalFormatting
      @worksheet = worksheet
    end

    # The worksheet that owns this collection of tables
    # @return [Worksheet]
    attr_reader :worksheet

    # serialize the conditional formattings
    def to_xml_string(str = "")
      return if empty?
      each { |item| item.to_xml_string(str) }
    end
  end

end

