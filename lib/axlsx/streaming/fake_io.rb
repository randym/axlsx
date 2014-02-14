module Axlsx

  module Streaming

    class FakeIO

      def initialize(&block)
        @block = block
        @pos = 0
      end

      def tell
        @pos
      end

      def pos
        @pos
      end

      def seek
        throw :fit
      end

      def pos=
        throw :fit
      end

      def to_s
        throw :fit
      end

      def <<(x)
        return if x.nil?
        throw "bad class #{x.class}" unless x.class == String
        @pos += x.bytesize
        @block.call(x.to_s)
      end

      def close
        nil
      end

    end

  end

end
