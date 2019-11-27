# frozen_string_literal: true
module Axlsx
  # This module allows us to define a list of symbols defining which
  # attributes will be serialized for a class.
  module SerializedAttributes

    # Extend with class methods
    def self.included(base)
      base.send :extend, ClassMethods
    end

    # class methods applied to all includers
    module ClassMethods

      # This is the method to be used in inheriting classes to specify
      # which of the instance values are serializable
      def serializable_attributes(*symbols)
        @xml_attributes = symbols
      end

      # a reader for those attributes
      attr_reader :xml_attributes

      # This helper registers the attributes that will be formatted as elements.
      def serializable_element_attributes(*symbols)
        @xml_element_attributes = symbols
      end

      # attr reader for element attributes
      attr_reader :xml_element_attributes
    end

    # creates a XML tag with serialized attributes
    # @see SerializedAttributes#serialized_attributes
    def serialized_tag(tagname, str, additional_attributes = {}, &block)
      str << "<#{tagname} "
      serialized_attributes(str, additional_attributes)
      if block_given?
        str << '>'
        yield
        str << "</#{tagname}>"
      else
        str << '/>'
      end
    end

    # serializes the instance values of the defining object based on the
    # list of serializable attributes.
    # @param [String] str The string instance to append this
    # serialization to.
    # @param [Hash] additional_attributes An option key value hash for
    # defining values that are not serializable attributes list.
    def serialized_attributes(str = String.new, additional_attributes = {})
      attributes = declared_attributes.merge! additional_attributes
      attributes.each do |key, value|
        str << "#{Axlsx.camel(key, false)}=\"#{Axlsx.camel(Axlsx.booleanize(value), false)}\" "
      end
      str
    end

    # A hash of instance variables that have been declared with
    # seraialized_attributes and are not nil.
    # This requires ruby 1.9.3 or higher
    def declared_attributes
      instance_values.select do |key, value|
        value != nil && self.class.xml_attributes.include?(key.to_sym)
      end
    end

    # serialized instance values at text nodes on a camelized element of the
    # attribute name. You may pass in a block for evaluation against non nil
    # values. We use an array for element attributes becuase misordering will
    # break the xml and 1.8.7 does not support ordered hashes.
    # @param [String] str The string instance to which serialized data is appended
    # @param [Array] additional_attributes An array of additional attribute names.
    # @return [String] The serialized output.
    def serialized_element_attributes(str = String.new, additional_attributes = [], &block)
      attrs = self.class.xml_element_attributes + additional_attributes
      values = instance_values
      attrs.each do |attribute_name|
        value = values[attribute_name.to_s]
        next if value.nil?
        value = yield value if block_given?
        element_name = Axlsx.camel(attribute_name, false)
        str << "<#{element_name}>#{value}</#{element_name}>"
      end
      str
    end
  end
end
