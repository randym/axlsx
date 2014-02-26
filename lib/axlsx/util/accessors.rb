module Axlsx
  # This module defines some of the more common validating attribute
  # accessors that we use in Axlsx
  #
  # When this module is included in your class you can simply call
  #
  # string_attr_access :foo
  #
  # To generate a new, validating set of accessors for foo.
  module Accessors
    def self.included(base)
      base.send :extend, ClassMethods
    end

    # Defines the class level xxx_attr_accessor methods
    module ClassMethods

      # Creates one or more string validated attr_accessors
      # @param [Array] symbols An array of symbols representing the
      # names of the attributes you will add to your class.
      def string_attr_accessor(*symbols)
        validated_attr_accessor(symbols, :validate_string)
      end


      # Creates one or more usigned integer attr_accessors
      # @param [Array] symbols An array of symbols representing the
      # names of the attributes you will add to your class
      def unsigned_int_attr_accessor(*symbols)
        validated_attr_accessor(symbols, :validate_unsigned_int)
      end

      # Creates one or more float (double?) attr_accessors
      # @param [Array] symbols An array of symbols representing the
      # names of the attributes you will add to your class
      def float_attr_accessor(*symbols)
        validated_attr_accessor(symbols, :validate_float)
      end

      # Creates on or more boolean validated attr_accessors
      # @param [Array] symbols An array of symbols representing the
      # names of the attributes you will add to your class.
      def boolean_attr_accessor(*symbols)
        validated_attr_accessor(symbols, :validate_boolean)
      end

      # Template for defining validated write accessors
      SETTER = "def %s=(value) Axlsx::%s(value); @%s = value; end".freeze

      # Creates the reader and writer access methods
      # @param [Array] symbols The names of the attributes to create
      # @param [String] validator The axlsx validation method to use when
      # validating assignation.
      # @see lib/axlsx/util/validators.rb 
      def validated_attr_accessor(symbols, validator)
        symbols.each do |symbol|
          attr_reader symbol
          module_eval(SETTER % [symbol, validator, symbol], __FILE__, __LINE__)
        end
      end
    end
  end
end

