module Axlsx
  class RichText < SimpleTypedList
    def initialize(text = nil, options={})
      super(RichTextRun)
      add_run(text, options) unless text.nil?
      yield self if block_given?
    end

    attr_reader :cell
    
    def cell=(cell)
      @cell = cell
      each { |run| run.cell = cell }
    end
    
    def autowidth
      widtharray = [0] # Are arrays the best way of solving this problem?
      each { |run| run.autowidth(widtharray) }
      widtharray.max
    end

    def add_run(text, options={})
      self << RichTextRun.new(text, options)
    end
    
    def runs
      self
    end

    def to_xml_string(str='')
      each{ |run| run.to_xml_string(str) }
      str
    end
  end
end
