module Axlsx

  # A simple, self serializing class for storing TextRuns
  class RichText < SimpleTypedList

    # creates a new RichText collection
    # @param [String] text -optional The text to use in creating the first RichTextRun
    # @param [Array] text -optional An array of Hashes to use to create multiple RichTextRuns
    # @param [Object] options -optional The options to use in creating the first RichTextRun
    # @yield [RichText] self
    def initialize(text = nil, options={})
      super(RichTextRun)
      unless text.nil?
        if text.is_a?(Array)
          text.each do |run_hash|
            add_run(run_hash[:text], run_hash[:options] || {})
          end
        else
          add_run(text, options)
        end
      end

      yield self if block_given?
    end

    # The cell that owns this RichText collection
    attr_reader :cell

    # Assign the cell for this RichText collection
    # @param [Cell] cell The cell which all RichTextRuns in the collection will belong to
    def cell=(cell)
      @cell = cell
      each { |run| run.cell = cell }
    end

    # Calculates the longest autowidth of the RichTextRuns in this collection
    # @return [Number]
    def autowidth
      widtharray = [0] # Are arrays the best way of solving this problem?
      each { |run| run.autowidth(widtharray) }
      widtharray.max
    end

    # Creates and adds a RichTextRun to this collectino
    # @param [String] text The text to use in creating a new RichTextRun
    # @param [Object] options The options to use in creating the new RichTextRun
    def add_run(text, options={})
      self << RichTextRun.new(text, options)
    end

    # The RichTextRuns we own
    # @return [RichText]
    def runs
      self
    end

    # renders the RichTextRuns in this collection
    # @param [String] str
    # @return [String]
    def to_xml_string(str='')
      each{ |run| run.to_xml_string(str) }
      str
    end
  end
end
