# encoding: UTF-8
module Axlsx
  # The picture locking class defines the locking properties for pictures in your workbook. 
  class PictureLocking
    
    
    attr_reader :noGrp
    attr_reader :noSelect
    attr_reader :noRot
    attr_reader :noChangeAspect
    attr_reader :noMove
    attr_reader :noResize
    attr_reader :noEditPoints
    attr_reader :noAdjustHandles
    attr_reader :noChangeArrowheads
    attr_reader :noChangeShapeType

    # Creates a new PictureLocking object
    # @option options [Boolean] noGrp
    # @option options [Boolean] noSelect
    # @option options [Boolean] noRot
    # @option options [Boolean] noChangeAspect
    # @option options [Boolean] noMove
    # @option options [Boolean] noResize
    # @option options [Boolean] noEditPoints
    # @option options [Boolean] noAdjustHandles
    # @option options [Boolean] noChangeArrowheads
    # @option options [Boolean] noChangeShapeType
    def initialize(options={})
      @noChangeAspect = true
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
    end        

    # @see noGrp
    def noGrp=(v) Axlsx::validate_boolean v; @noGrp = v end    

    # @see noSelect
    def noSelect=(v) Axlsx::validate_boolean v; @noSelect = v end    

    # @see noRot
    def noRot=(v) Axlsx::validate_boolean v; @noRot = v end    

    # @see noChangeAspect
    def noChangeAspect=(v) Axlsx::validate_boolean v; @noChangeAspect = v end    

    # @see noMove
    def noMove=(v) Axlsx::validate_boolean v; @noMove = v end    

    # @see noResize
    def noResize=(v) Axlsx::validate_boolean v; @noResize = v end    

    # @see noEditPoints
    def noEditPoints=(v) Axlsx::validate_boolean v; @noEditPoints = v end    

    # @see noAdjustHandles
    def noAdjustHandles=(v) Axlsx::validate_boolean v; @noAdjustHandles = v end    

    # @see noChangeArrowheads
    def noChangeArrowheads=(v) Axlsx::validate_boolean v; @noChangeArrowheads = v end    

    # @see noChangeShapeType
    def noChangeShapeType=(v) Axlsx::validate_boolean v; @noChangeShapeType = v end    

    # Serializes the picture locking
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml[:a].picLocks(self.instance_values)      
    end
  end
end
