module Axlsx

  class RowContainer

    def initialize(row_enumerator)
      @row_enumerator = row_enumerator
      @rows = SimpleTypedList.new Row
    end

    def each(&block)
      rows = @row_enumerator ? @row_enumerator : @rows
      rows.each(&block)
    end

    def each_with_index(&block)
      rows = @row_enumerator ? @row_enumerator : @rows
      rows.each_with_index(&block)
    end

    def index(row)
      @rows.index(row)
    end

    def <<(row)
      @rows << row
    end

    def [](index)
      @rows[index]
    end

    def empty?
      @rows.empty?
    end

    def first
      @rows.first
    end

    def last
      @rows.last
    end

    def size
      @rows.size
    end

    def transpose(&block)
      @rows.transpose(&block)
    end

    def flatten
      @rows.flatten
    end

  end

end
