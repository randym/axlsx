require 'zip'

module Axlsx

  module Streaming

    class OutputStream < Zip::OutputStream

      class << self
        def open(io)
          zos = new(io)
          yield zos
        ensure
          zos.close if zos
        end
      end

      def initialize(io)
        super io, true
        @output_stream = io
      end

      def put_next_entry(zip_entry)
        init_next_entry(zip_entry)
        @current_entry = zip_entry
      end

      def finalize_current_entry
        if @current_entry
          entry = @current_entry
          super
          write_local_footer(entry)
        end
      end

      def write_local_footer(entry)
        @output_stream << [ 0x08074b50, entry.crc, entry.compressed_size, entry.size].pack('VVVV')
      end

      def update_local_headers
        nil
      end

    end

  end

end
