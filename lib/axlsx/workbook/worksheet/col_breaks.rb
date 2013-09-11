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
      @list << Break.new(options.merge(:max => 1048575, :man => true))
      last
    end

    # Serialize the collection to xml
    # @param [String] str The string to append this lists xml to.
    # <colBreaks count="1" manualBreakCount="1">
    # <brk id="3" max="1048575" man="1"/>
    # </colBreaks>
    def to_xml_string(str='')
      return if empty?
      str << '<colBreaks count="' << @list.size.to_s << '" manualBreakCount="' << @list.size.to_s << '">'
      each { |brk| brk.to_xml_string(str) }
      str << '</colBreaks>'
    end
  end
end
