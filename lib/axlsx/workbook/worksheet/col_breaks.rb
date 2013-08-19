module Axlsx

  class ColBreaks < SimpleTypedList

    def initialize
      super Break
    end

    def add_break(options)
      options.merge max: 1048575, man: true
      @list << Break.new(options)
    end

    def to_xml_string(str='')
      return if empty?
      str << '<colBreaks count="' << @list.size << '" manualBreakCount="' << @list.size << '">'
      each { |brk| brk.to_xml_string(str) }
      str << '</rowBreaks>'
    end
  end
end
