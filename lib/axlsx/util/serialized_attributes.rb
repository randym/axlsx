module Axlsx
  module SerializedAttributes

    def self.included(base)
      base.send :extend, ClassMethods
    end

    module ClassMethods
      def serializable_attributes(*symbols)
        @@serializable_attributes = symbols
      end
    end

    def serialized_attributes(str = '')
      serializable_values = instance_values.select do |key, value|
        self.class.serializable_attributes.include?(key.to_sym)
      end
      serializable_values.each do |key, value|
        str << "#{Axlsx.camel(key, false)}='#{value}' "
      end
      str
    end
  end
end
