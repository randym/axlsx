# encoding: UTF-8
module Axlsx

  # The SheetProtection object manages worksheet protection options per sheet.
  class SheetProtection

    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes
    include Axlsx::Accessors

    # Creates a new SheetProtection instance
    # @option options [Boolean] sheet @see SheetProtection#sheet
    # @option options [Boolean] objects @see SheetProtection#objects
    # @option options [Boolean] scenarios @see SheetProtection#scenarios
    # @option options [Boolean] format_cells @see SheetProtection#objects
    # @option options [Boolean] format_columns @see SheetProtection#format_columns
    # @option options [Boolean] format_rows @see SheetProtection#format_rows
    # @option options [Boolean] insert_columns @see SheetProtection#insert_columns
    # @option options [Boolean] insert_rows @see SheetProtection#insert_rows
    # @option options [Boolean] insert_hyperlinks @see SheetProtection#insert_hyperlinks
    # @option options [Boolean] delete_columns @see SheetProtection#delete_columns
    # @option options [Boolean] delete_rows @see SheetProtection#delete_rows
    # @option options [Boolean] select_locked_cells @see SheetProtection#select_locked_cells
    # @option options [Boolean] sort @see SheetProtection#sort
    # @option options [Boolean] auto_filter @see SheetProtection#auto_filter
    # @option options [Boolean] pivot_tables @see SheetProtection#pivot_tables
    # @option options [Boolean] select_unlocked_cells @see SheetProtection#select_unlocked_cells
    # @option options [String] password. The password required for unlocking. @see SheetProtection#password=
    def initialize(options={})
      @objects = @scenarios = @select_locked_cells = @select_unlocked_cells = false
      @sheet = @format_cells = @format_rows = @format_columns = @insert_columns = @insert_rows = @insert_hyperlinks = @delete_columns = @delete_rows = @sort = @auto_filter = @pivot_tables = true
      @password = nil
      parse_options options
    end

    boolean_attr_accessor :sheet, :objects, :scenarios, :format_cells, :format_columns, :format_rows,
      :insert_columns, :insert_rows, :insert_hyperlinks, :delete_columns, :delete_rows,
      :select_locked_cells, :sort, :auto_filter, :pivot_tables, :select_unlocked_cells

    serializable_attributes :sheet, :objects, :scenarios, :format_cells, :format_columns, :format_rows,
      :insert_columns, :insert_rows, :insert_hyperlinks, :delete_columns, :delete_rows,
      :select_locked_cells, :sort, :auto_filter, :pivot_tables, :select_unlocked_cells, :salt, :password

    # Specifies the salt which was prepended to the user-supplied password before it was hashed using the hashing algorithm
    # @return [String]
    attr_reader :salt_value

    # Password hash
    # @return [String]
    # default nil
    attr_reader :password

    # This block is intended to implement the salt_value, hash_value and spin count as per the ECMA-376 standard. 
    # However, it does not seem to actually work in EXCEL - instead they are using their old retro algorithm shown below 
    # defined in the transitional portion of the speck. I am leaving this code in in the hope that someday Ill be able to 
    # figure out why it does not work, and if Excel even supports it.
    #     def propper_password=(v)
    #       @algorithm_name = v == nil ? nil : 'SHA-1'
    #       @salt_value = @spin_count = @hash_value = v if v == nil
    #       return if v == nil
    #       require 'digest/sha1'
    #       @spin_count = 10000
    #       @salt_value = Digest::SHA1.hexdigest(rand(36**8).to_s(36))
    #       @spin_count.times do |count|
    #         @hash_value = Digest::SHA1.hexdigest((@hash_value ||= (@salt_value + v.to_s)) + Array(count).pack('V'))
    #       end
    #     end

    # encodes password for protection locking
    def password=(v)
      return if v == nil
      @password = create_password_hash(v)
    end

    # Serialize the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<sheetProtection '
      serialized_attributes str
      str << '/>'
    end

    private
    # Creates a password hash for a given password
    # @return [String]
    def create_password_hash(password)
      encoded_password = encode_password(password)

      password_as_hex = [encoded_password].pack("v")
      password_as_string = password_as_hex.unpack("H*").first.upcase

      password_as_string[2..3] + password_as_string[0..1]
    end

    # Encodes a given password
    # Based on the algorithm provided by Daniel Rentz of OpenOffice.
    # http://www.openoffice.org/sc/excelfileformat.pdf, Revision 1.42, page 115 (21.05.2012)
    # @return [String]
    def encode_password(password)
      i = 0
      chars = password.split(//)
      count = chars.size

      chars.collect! do |char|
        i += 1
        char     = char.unpack('c')[0] << i #ord << i
        low_15   = char & 0x7fff
        high_15  = char & 0x7fff << 15
        high_15  = high_15 >> 15
        char     = low_15 | high_15
      end

      encoded_password  = 0x0000
      chars.each { |c| encoded_password ^= c }
      encoded_password ^= count
      encoded_password ^= 0xCE4B
    end
  end
end
