module Axlsx
  # The Protected Range class represents a set of cells in the worksheet
  # @note the recommended way to manage protected ranges with via Worksheet#protect_range
  # @see Worksheet#protect_range
  class ProtectedRange

    # The reference for the protected range
    # @return [String]
    attr_reader :sqref

    # The name of the protected range
    # @return [String]
    attr_reader :name

    # Initializes a new protected range object
    # @option [String] sqref The cell range reference to protect. This can be an absolute or a relateve range however, it only applies to the current sheet.
    # @option [String] name An optional name for the protected name.
    def initialize(options={})
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
      yield self if block_given?
    end
    
    # @see sqref
    def sqref=(v)
      Axlsx.validate_string(v)
      @sqref = v
    end

    # @see name
    def name=(v)
      Axlsx.validate_string(v)
      @name = v
    end
   
    # serializes the proteted range
    # @param [String] str if this string object is provided we append
    # our output to that object. Use this - it helps limit the number of
    # objects created during serialization
    def to_xml_string(str="")
      attrs = self.instance_values.reject{ |key, value| value == nil }
      str << '<protectedRange ' << attrs.map { |key, value| '' << key << '="' << value.to_s << '"' }.join(' ') << '/>'
    end 
  end
end
