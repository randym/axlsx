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
      def xml_attributes
        @xml_attributes
      end
    end

    # serializes the instance values of the defining object based on the 
    # list of serializable attributes.
    # @param [String] str The string instance to append this
    # serialization to.
    # @param [Hash] additional_attributes An option key value hash for
    # defining values that are not serializable attributes list.
    def serialized_attributes(str = '', additional_attributes = {})
      key_value_pairs = instance_values
      key_value_pairs.each do |key, value|
        key_value_pairs.delete(key) if value == nil
        key_value_pairs.delete(key) unless self.class.xml_attributes.include?(key.to_sym)
      end
      key_value_pairs.merge! additional_attributes

      key_value_pairs.each do |key, value|
        str << "#{Axlsx.camel(key, false)}=\"#{value}\" "
      end
      str
    end
  end
end
