# frozen_string_literal: true
module Axlsx

  #A collection of hyperlink objects for a worksheet
  class WorksheetHyperlinks < SimpleTypedList

    # Creates a new Hyperlinks collection
    # @param [Worksheet] worksheet the worksheet that owns these hyperlinks
    def initialize(worksheet)
      DataTypeValidator.validate "Hyperlinks.worksheet", [Worksheet], worksheet
      @worksheet = worksheet
      super WorksheetHyperlink
    end

    # Creates and adds a new hyperlink based on the options provided
    # @see WorksheetHyperlink#initialize
    # @return [WorksheetHyperlink]
    def add(options)
      self << WorksheetHyperlink.new(@worksheet, options)
      last
    end

    # The relationships required by this collection's hyperlinks
    # @return Array
    def relationships
      return [] if empty?
      map { |hyperlink| hyperlink.relationship }
    end

    # seralize the collection of hyperlinks
    # @return [String]
    def to_xml_string(str = String.new)
      return if empty?
      str << '<hyperlinks>'
      each { |hyperlink| hyperlink.to_xml_string(str) }
      str << '</hyperlinks>'
    end
  end
end
