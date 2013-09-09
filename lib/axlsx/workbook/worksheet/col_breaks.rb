module Axlsx

  # A collection of Brake objects.
  # Please do not use this class directly. Instead use 
  # Worksheet#add_break
  class ColBreaks < SimpleTypedList

    # Instantiates a new list restricted to Break types
    def initialize
      super Break
    end

    # A column break specific helper for adding a break.
    # @param [Hash] options A list of options to pass into the Break object
    # The max and man options are fixed, however any other valid option for 
    # Break will be passed to the created break object.
    # @see Break
    def add_break(options)
      options.merge :max => 1048575, :man => true
      @list << Break.new(options)
    end

    # Serialize the collection to xml
    # @param [String] str The string to append this lists xml to.
    def to_xml_string(str='')
      return if empty?
      str << '<colBreaks count="' << @list.size << '" manualBreakCount="' << @list.size << '">'
      each { |brk| brk.to_xml_string(str) }
      str << '</rowBreaks>'
    end
  end
end
