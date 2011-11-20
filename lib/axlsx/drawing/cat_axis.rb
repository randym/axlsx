module Axlsx
  #A CatAxis object defines a chart category axis
  class CatAxis < Axis
    # From the docs: This element specifies that this axis is a date or text axis based on the data that is used for the axis labels, not a specific choice.
    # @return [Boolean]
    attr_accessor :auto

    # specifies how the perpendicular axis is crossed
    # must be one of [:ctr, :l, :r]
    # @return [Symbol]
    attr_accessor :lblAlgn 

    # The offset of the labels
    # must be between a string between 0 and 1000
    # @return [Integer]
    attr_accessor :lblOffset

    # regex for validating label offset
    LBL_OFFSET_REGEX = /0*(([0-9])|([1-9][0-9])|([1-9][0-9][0-9])|1000)%/

    # Creates a new CatAxis object
    # @param [Integer] axId the id of this axis
    # @param [Integer] crossAx the id of the perpendicular axis
    # @option options [Symbol] axPos
    # @option options [Symbol] tickLblPos
    # @option options [Symbol] crosses
    # @option options [Boolean] auto
    # @option options [Symbol] lblAlgn
    # @option options [Integer] lblOffset    
    def initialize(axId, crossAx, options={})
      super(axId, crossAx, options)
      self.auto = true
      self.lblAlgn = :ctr
      self.lblOffset = "100%"
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
    end 

    def auto=(v) Axlsx::validate_boolean(v); @auto = v; end
    def lblAlgn=(v) RestrictionValidator.validate "#{self.class}.lblAlgn", [:ctr, :l, :r], v; @lblAlgn = v; end
    def lblOffset=(v) RegexValidator.validate "#{self.class}.lblOffset", LBL_OFFSET_REGEX, v; @lblOffset = v; end

    # Serializes the category axis
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml.send('c:catAx') {
        super(xml)
        xml.send('c:auto', :val=>@auto)
        xml.send('c:lblAlgn', :val=>@lblAlgn)
        xml.send('c:lblOffset', :val=>@lblOffset)                 
      }
    end
  end
  

end
