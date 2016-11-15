# encoding: UTF-8
module Axlsx

  # A SimpleTypedList is a type restrictive collection that allows some of the methods from Array and supports basic xml serialization.
  # @private
  class SimpleTypedList
    # Creats a new typed list
    # @param [Array, Class] type An array of Class objects or a single Class object
    # @param [String] serialize_as The tag name to use in serialization
    # @raise [ArgumentError] if all members of type are not Class objects
    def initialize type, serialize_as=nil, start_size = 0
      if type.is_a? Array
        type.each { |item| raise ArgumentError, "All members of type must be Class objects" unless item.is_a? Class }
        @allowed_types = type
      else
        raise ArgumentError, "Type must be a Class object or array of Class objects" unless type.is_a? Class
        @allowed_types = [type]
      end
      @serialize_as = serialize_as unless serialize_as.nil?
      @list = Array.new(start_size)
    end

    # The class constants of allowed types
    # @return [Array]
    attr_reader :allowed_types

    # The index below which items cannot be removed
    # @return [Integer]
    def locked_at
      defined?(@locked_at) ? @locked_at : nil
    end

    # The tag name to use when serializing this object
    # by default the parent node for all items in the list is the classname of the first allowed type with the first letter in lowercase.
    # @return [String]
    attr_reader :serialize_as

    # Transposes the list (without blowing up like ruby does)
    # any non populated cell in the matrix will be a nil value
    def transpose
      return @list.clone if @list.size == 0
      row_count = @list.size
      max_column_count = @list.map{|row| row.cells.size}.max
      result = Array.new(max_column_count) { Array.new(row_count) }
      # yes, I know it is silly, but that warning is really annoying
      row_count.times do |row_index|
         max_column_count.times do |column_index|
          datum = if @list[row_index].cells.size >= max_column_count
                    @list[row_index].cells[column_index]
                  elsif block_given?
                    yield(column_index, row_index)
                  end
          result[column_index][row_index] = datum
        end
      end
      result
    end
    
    # Lock this list at the current size
    # @return [self]
    def lock
      @locked_at = @list.size
      self
    end

    # Unlock the list
    # @return [self]
    def unlock
      @locked_at = nil
      self
    end
    
    def to_ary
      @list
    end

    alias :to_a :to_ary

    # join operator
    # @param [Array] v the array to join
    # @raise [ArgumentError] if any of the values being joined are not
    # one of the allowed types
    # @return [SimpleTypedList]
    def +(v)
      v.each do |item| 
        DataTypeValidator.validate :SimpleTypedList_plus, @allowed_types, item
        @list << item 
      end
    end

    # Concat operator
    # @param [Any] v the data to be added
    # @raise [ArgumentError] if the value being added is not one fo the allowed types
    # @return [Integer] returns the index of the item added.
    def <<(v)
      DataTypeValidator.validate :SimpleTypedList_push, @allowed_types, v
      @list << v
      @list.size - 1
    end 
    
    alias :push :<<
    

    # delete the item from the list
    # @param [Any] v The item to be deleted.
    # @raise [ArgumentError] if the item's index is protected by locking
    # @return [Any] The item deleted
    def delete(v)
      return unless include? v
      raise ArgumentError, "Item is protected and cannot be deleted" if protected? index(v)
      @list.delete v
    end

    # delete the item from the list at the index position provided
    # @raise [ArgumentError] if the index is protected by locking
    # @return [Any] The item deleted
    def delete_at(index)
      @list[index]
      raise ArgumentError, "Item is protected and cannot be deleted" if protected? index
      @list.delete_at index
    end

    # positional assignment. Adds the item at the index specified
    # @param [Integer] index
    # @param [Any] v
    # @raise [ArgumentError] if the index is protected by locking
    # @raise [ArgumentError] if the item is not one of the allowed types
    def []=(index, v)
      DataTypeValidator.validate :SimpleTypedList_insert, @allowed_types, v
      raise ArgumentError, "Item is protected and cannot be changed" if protected? index
      @list[index] = v
      v
    end

    # inserts an item at the index specfied
    # @param [Integer] index
    # @param [Any] v
    # @raise [ArgumentError] if the index is protected by locking
    # @raise [ArgumentError] if the index is not one of the allowed types
    def insert(index, v)
      DataTypeValidator.validate :SimpleTypedList_insert, @allowed_types, v
      raise ArgumentError, "Item is protected and cannot be changed" if protected? index
      @list.insert(index, v)
      v
    end

    # determines if the index is protected
    # @param [Integer] index
    def protected? index
      return false unless locked_at.is_a? Integer
      index < locked_at
    end

    DESTRUCTIVE = ['replace', 'insert', 'collect!', 'map!', 'pop', 'delete_if',
                   'reverse!', 'shift', 'shuffle!', 'slice!', 'sort!', 'uniq!',
                   'unshift', 'zip', 'flatten!', 'fill', 'drop', 'drop_while',
                   'delete_if', 'clear']
    DELEGATES = Array.instance_methods - self.instance_methods - DESTRUCTIVE

    DELEGATES.each do |method|
      class_eval %{
        def #{method}(*args, &block)
          @list.send(:#{method}, *args, &block)
        end
      }
    end
                   
    def to_xml_string(str = '')
      classname = @allowed_types[0].name.split('::').last
      el_name = serialize_as.to_s || (classname[0,1].downcase + classname[1..-1])
      str << ('<' << el_name << ' count="' << size.to_s << '">')
      each { |item| item.to_xml_string(str) }
      str << ('</' << el_name << '>')
    end

  end


end
