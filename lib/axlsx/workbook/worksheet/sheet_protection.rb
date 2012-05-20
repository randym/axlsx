# encoding: UTF-8
module Axlsx

  # The SheetProtection object manages worksheet protection options per sheet.
  class SheetProtection

    # Specifies the specific cryptographic hashing algorithm which shall be used along
    # with the salt attribute and input password in order to compute the hash value.
    # This value is automatically set to 'SHA-1' when password= is called.
    # @note only SHA-1 is supported.
    # @return [String]
    attr_reader :algorithm_name


    # If 1 or true then AutoFilters should not be allowed to operate when the sheet is protected.
    # If 0 or false then AutoFilters should be allowed to operate when the sheet is protected.
    # @return [Boolean]
    # @default true 
    attr_reader :auto_filter

    # If 1 or true then deleting columns should not be allowed when the sheet is protected. 
    # If 0 or false then deleting columns should be allowed when the sheet is protected.
    # @return [Boolean]
    # @default true
    attr_reader :delete_columns

    # If 1 or true then deleting rows should not be allowed when the sheet is protected.
    # If 0 or false then deleting rows should be allowed when the sheet is protected.
    # @return [Boolean]
    # @default true
    attr_reader :delete_rows

    # If 1 or true then formatting cells should not be allowed when the sheet is protected.
    # If 0 or false then formatting cells should be allowed when the sheet is protected.
    # @return [Boolean]
    # @default true
    attr_reader :format_cells

    # If 1 or true then formatting columns should not be allowed when the sheet is protected.
    # If 0 or false then formatting columns should be allowed when the sheet is protected.
    # @return [Boolean]
    # @default true
    attr_reader :format_columns

    # If 1 or true then formatting rows should not be allowed when the sheet is protected. 
    # If 0 or false then formatting rows should be allowed when the sheet is protected.
    # @return [Boolean]
    # @default true
    attr_reader :format_rows

    # Specifies the hash value for the password required to edit this worksheet. 
    # @return [String]
    attr_reader :hash_value

    # If 1 or true then inserting columns should not be allowed when the sheet is protected. 
    # If 0 or false then inserting columns should be allowed when the sheet is protected.
    # @return [Boolean]
    # @default true
    attr_reader :insert_columns

    # If 1 or true then inserting hyperlinks should not be allowed when the sheet is protected. 
    # If 0 or false then inserting hyperlinks should be allowed when the sheet is protected.
    # @return [Boolean]
    # @default true
    attr_reader :insert_hyperlinks

    # If 1 or true then inserting rows should not be allowed when the sheet is protected. 
    # If 0 or false then inserting rows should be allowed when the sheet is protected.
    # @return [Boolean]
    # @default true
    attr_reader :insert_rows

    # If 1 or true then editing of objects should not be allowed when the sheet is protected. 
    # If 0 or false then objects are allowed to be edited when the sheet is protected.
    # @return [Boolean]
    # @default false
    attr_reader :objects

    # If 1 or true then PivotTables should not be allowed to operate when the sheet is protected.
    # If 0 or false then PivotTables should be allowed to operate when the sheet is protected.
    # @return [Boolean]
    # @default true
    attr_reader :pivot_tables

    # Specifies the salt which was prepended to the user-supplied password before it was hashed using the hashing algorithm
    # @return [String]
    attr_reader :salt_value

    # If 1 or true then Scenarios should not be edited when the sheet is protected.
    # If 0 or false then Scenarios are allowed to be edited when the sheet is protected.
    # @return [Boolean]
    # @default false
    attr_reader :scenarios

    # If 1 or true then selection of locked cells should not be allowed when the sheet is protected.
    # If 0 or false then selection of locked cells should be allowed when the sheet is protected.
    # @return [Boolean]
    # @default false
    attr_reader :select_locked_cells

    # If 1 or true then selection of unlocked cells should not be allowed when the sheet is protected.
    # If 0 or false then selection of unlocked cells should be allowed when the sheet is protected.
    # @return [Boolean]
    # @default false
    attr_reader :select_unlocked_cells

    # If 1 or true then the sheet is protected. 
    # If 0 or false then the sheet is not protected. 
    # @return [Boolean]
    # @default true
    attr_reader :sheet

    # If 1 or true then sorting should not be allowed when the sheet is protected.
    # If 0 or false then sorting should be allowed when the sheet is protected.
    # @return [Boolean]
    # @default true
    attr_reader :sort

    # Specifies the number of times the hashing function shall be iteratively run
    # @return [Integer]
    # @default 10000
    attr_reader :spin_count

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
    # @option options [Boolean] objects @see SheetProtection#objects
    def initialize(options={})
      @objects = @scenarios = @select_locked_cells = @select_unlocked_cells = false
      @sheet = @format_cells = @format_rows = @format_columns = @insert_columns = @insert_rows = @insert_hyperlinks = @delete_columns = @delete_rows = @sort = @auto_filter = @pivot_tables = true
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
    end

    [:sheet, :objects, :scenarios, :select_locked_cells, :sort,
     :select_unlocked_cells, :format_cells, :format_rows, :format_columns,
     :insert_columns, :insert_rows, :insert_hyperlinks, :delete_columns, 
     :delete_rows, :auto_filter, :pivot_tables].each do |f_name|
       define_method "#{f_name.to_s}=".to_sym do |v| 
         Axlsx::validate_boolean(v)
         instance_variable_set "@#{f_name.to_s}".to_sym, v
       end
     end

     def password=(v)
       @algorithm_name = v == nil ? nil : 'SHA-1'
       @salt_value = @spin_count = @hash_value = v if v == nil
       return if v == nil
       require 'digest/sha1'
       @spin_count = 10000
       @salt_value = Digest::SHA1.hexdigest(rand(36**8).to_s(36))
       @hash_value = nil
       @spin_count.times do |count|
         @hash_value = Digest::SHA1.hexdigest((@hash_value || (@salt_value + v.to_s)) + count.to_s.bytes.to_a.pack('l'))
       end
     end

     def to_xml_string(str = '')
       str << '<sheetProtection '
       str << instance_values.map{ |k,v| k.gsub(/_(.)/){ $1.upcase } << %{="#{v.to_s}"} }.join(' ')
       str << '/>'
     end 
  end
end
