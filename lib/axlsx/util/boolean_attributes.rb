module Axlsx
  module BooleanAttributes
    def BooleanAttributes.included(mod)
      mod::BOOLEAN_ATTRIBUTES.each do |attr|
        class_eval %{
        # The #{attr} attribute reader
        # @return [Boolean]
        attr_reader :#{attr} 

        # The #{attr} writer
        # @param [Boolean] value The value to assign to #{attr}
        # @return [Boolean]
        def #{attr}=(value)
          Axlsx::validate_boolean(value)
          @#{attr} = value
        end
        }
      end
    end 
  end

  module StringAttributes
    def StringAttributes.included(mod)
      mod::STRING_ATTRIBUTES.each do |attr|
        class_eval %{
        # The #{attr} attribute reader
        # @return [String]
        attr_reader :#{attr}

        # The #{attr} writer
        # @param [String] value The value to assign to #{attr}
        # @return [String]
        def #{attr}=(value)
          Axlsx::validate_string(value)
          @#{attr}= value
        end
        }

      end
    end
  end

end
