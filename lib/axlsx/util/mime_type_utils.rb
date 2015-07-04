module Axlsx
  # This module defines some utils related with mime type detection
  module MimeTypeUtils
    # Detect a file mime type
    # @param [String] v File path
    # @return [String] File mime type
    def self.get_mime_type(v)
      MimeMagic.by_magic(File.open(v)).to_s
    end
  end
end
