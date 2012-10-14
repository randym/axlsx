module Axlsx
  module SerializedAttributes

    def self.included(base)
      base.send :extend, ClassMethods
    end

    module ClassMethods

      def serializable_attributes(*symbols)
        @xml_attributes = symbols
      end
      def xml_attributes
        @xml_attributes
      end
    end

    def serialized_attributes(str = '', additional_attributes = {})
      key_value_pairs = instance_values
      key_value_pairs.each do |key, value|
        key_value_pairs.delete(key) if value == nil
        key_value_pairs.delete(key) unless self.class.xml_attributes.include?(key.to_sym)
      end
      key_value_pairs.merge! additional_attributes

      key_value_pairs.each do |key, value|
        str << "#{Axlsx.camel(key, false)}='#{value}' "
      end
      str
    end
  end
end
