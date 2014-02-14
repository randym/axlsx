module Axlsx

  module Streaming

    class ZipBody

      def initialize(parts, timestamp)
        @parts = parts
        @timestamp = timestamp
      end

      def each(&block)
        OutputStream.open( FakeIO.new(&block) ) do |zip|
          @parts.each do |part|
            next if part[:doc].nil?
            zip.put_next_entry(zip_entry_for_part(part))
            entry = part[:doc]
            case entry
            when String
              zip << entry
            when Enumerator
              entry.each { |line| zip << line }
            end
          end
        end
      end

      private

      def zip_entry_for_part(part)
        timestamp = Zip::DOSTime.at(@timestamp)
        zip_entry = Zip::Entry.new("", part[:entry], "", "", 0, 0, Zip::Entry::DEFLATED, 0, timestamp)
        #THIS IS THE MAGIC, tells zip to look after data for size, crc
        zip_entry.gp_flags |= 0x0008
        zip_entry
      end

    end

  end

end

