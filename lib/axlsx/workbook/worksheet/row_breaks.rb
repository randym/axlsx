module Axlsx

  # A collection of break objects that define row breaks (page breaks) for printing and preview

  class RowBreaks < SimpleTypedList

    def initialize
      super Break
    end

    def add_break(options)
      # force feed the excel default
      options.merge max: 16383, man: true
      @list << Break.new(options)
      last
    end

    def to_xml_string(str='')
      return if empty?
      str << '<rowBreaks count="' << @list.size << '" manualBreakCount="' << @list.size << '">'
      each { |brk| brk.to_xml_string(str) }
      str << '</rowBreaks>'
    end
  end
end
