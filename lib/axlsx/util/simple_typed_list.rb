# encoding: UTF-8
module Axlsx
  # A SimpleTypedList is a type restrictive collection that allows some of the methods from Array and supports basic xml serialization.
  # @private
  class SimpleTypedList
    # The class constants of allowed types
    # @return [Array]
    attr_reader :allowed_types

    # The index below which items cannot be removed
    # @return [Integer]
    attr_reader :locked_at

    # The tag name to use when serializing this object
    # by default the parent node for all items in the list is the classname of the first allowed type with the first letter in lowercase.
    # @return [String]
    attr_reader :serialize_as

    # Creats a new typed list
    # @param [Array, Class] type An array of Class objects or a single Class object
    # @param [String] serialize_as The tag name to use in serialization
    # @raise [ArgumentError] if all members of type are not Class objects
    def initialize type, serialize_as=nil
      if type.is_a? Array
        type.each { |item| raise ArgumentError, "All members of type must be Class objects" unless item.is_a? Class }
        @allowed_types = type
      else
        raise ArgumentError, "Type must be a Class object or array of Class objects" unless type.is_a? Class
        @allowed_types = [type]
      end
      @list = []
      @locked_at = nil
      @serialize_as = serialize_as
    end

    # Lock this list at the current size
    # @return [self]
    def lock
      @locked_at = @list.size
      self
    end

    def to_ary
      @list
    end

    alias :to_a :to_ary

    # Unlock the list
    # @return [self]
    def unlock
      @locked_at = nil
      self
    end

    # join operator
    # @param [Array] v the array to join
    # @raise [ArgumentError] if any of the values being joined are not
    # one of the allowed types
    # @return [SimpleTypedList]
    def +(v)
      v.each do |item| 
        DataTypeValidator.validate "SimpleTypedList.+", @allowed_types, item
        @list << item 
      end
    end

    # Concat operator
    # @param [Any] v the data to be added
    # @raise [ArgumentError] if the value being added is not one fo the allowed types
    # @return [Integer] returns the index of the item added.
    def <<(v)
      DataTypeValidator.validate "SimpleTypedList.<<", @allowed_types, v
      @list << v
      @list.size - 1
    end
    alias :push :<<

    # delete the item from the list
    # @param [Any] v The item to be deleted.
    # @raise [ArgumentError] if the item's index is protected by locking
    # @return [Any] The item deleted
    def delete(v)
      return unless @list.include? v
      raise ArgumentError, "Item is protected and cannot be deleted" if protected? @list.index(v)
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
      DataTypeValidator.validate "SimpleTypedList.<<", @allowed_types, v
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
      DataTypeValidator.validate "SimpleTypedList.<<", @allowed_types, v
      raise ArgumentError, "Item is protected and cannot be changed" if protected? index
      @list.insert(index, v)
      v
    end

    # determines if the index is protected
    # @param [Integer] index
    def protected? index
      return false unless @locked_at.is_a? Fixnum
      index < @locked_at
    end

    # override the equality method so that this object can be compared to a simple array.
    # if this object's list is equal to the specifiec array, we return true.
    def ==(v)
      v == @list
    end
    # method_mission override to pass allowed methods to the list.
    # @note
    #  the following methods are not allowed
    #   :replace
    #   :insert
    #   :collect!
    #   :map!
    #   :pop
    #   :delete_if
    #   :reverse!
    #   :shift
    #   :shuffle!
    #   :slice!
    #   :sort!
    #   :uniq!
    #   :unshift
    #   :zip
    #   :flatten!
    #   :fill
    #   :drop
    #   :drop_while
    #   :delete_if
    #   :clear
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
      str << '<' << el_name << ' count="' << @list.size.to_s << '">'
      @list.each { |item| item.to_xml_string(str) }
      str << '</' << el_name << '>'
    end

  end


end
