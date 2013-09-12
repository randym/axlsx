module Axlsx

  # A collection of break objects that define row breaks (page breaks) for printing and preview

  class RowBreaks < SimpleTypedList

    def initialize
      super Break
    end

    # Adds a row break
    # @param [Hash] options The options for the break to be created.
    # max and man values are fixed.
    # @see Break
    def add_break(options)
      # force feed the excel default
      @list << Break.new(options.merge(:max => 16383, :man => true))
      last
    end
 
    # <rowBreaks count="3" manualBreakCount="3">
    # <brk id="1" max="16383" man="1"/>
    # <brk id="7" max="16383" man="1"/>
    # <brk id="13" max="16383" man="1"/>
    # </rowBreaks>
    def to_xml_string(str='')
      return if empty?
      str << '<rowBreaks count="' << @list.size.to_s << '" manualBreakCount="' << @list.size.to_s << '">'
      each { |brk| brk.to_xml_string(str) }
      str << '</rowBreaks>'
    end
  end
end
