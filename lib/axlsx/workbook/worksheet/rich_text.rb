module Axlsx
  class RichText
    def initialize(text = nil, options={})
       @runs = SimpleTypedList.new(RichTextRun)
       add_run(text, options) unless text.nil?
       yield self if block_given?
    end

    attr_reader :runs
    
    attr_reader :cell
    
    def cell=(cell)
      @cell = cell
      @runs.each { |run| run.cell = cell }
    end
    
    def autowidth
      widtharray = [0] # Are arrays the best way of solving this problem?
      @runs.each { |run| run.autowidth(widtharray) }
      widtharray.max
    end

    def add_run(text, options={})
      @runs << RichTextRun.new(text, options)
    end

    def to_xml_string(str='')
      runs.each{ |run| run.to_xml_string(str) }
      str
    end
  end
end
