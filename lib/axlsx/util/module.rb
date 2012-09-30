class Module
  def string_attr_accessor(*symbols)
    validated_attr_accessor(symbols, 'validate_string')
  end

  def boolean_attr_accessor(*symbols)
    validated_attr_accessor(symbols, 'validate_boolean')
  end

  SETTER = "def %s=(value) Axlsx::%s(value); @%s = value; end"
  def validated_attr_accessor(symbols, validator)
    symbols.each do |symbol|
      attr_reader symbol
      module_eval(SETTER % [symbol.to_s, validator, symbol.to_s], __FILE__, __LINE__)
    end
  end

end
