# encoding: UTF-8
module Axlsx
  # 3D attributes for a chart.
  class View3D

    # Validation for hPercent
    H_PERCENT_REGEX = /0*(([5-9])|([1-9][0-9])|([1-4][0-9][0-9])|500)%/
    
    # validation for depthPercent
    DEPTH_PERCENT_REGEX = /0*(([2-9][0-9])|([1-9][0-9][0-9])|(1[0-9][0-9][0-9])|2000)%/

    # x rotation for the chart 
    # must be between -90 and 90
    # @return [Integer]
    attr_reader :rotX
    
    # height of chart as % of chart
    # must be between 5% and 500%
    # @return [String]
    attr_reader :hPercent
    
    # y rotation for the chart
    # must be between 0 and 360
    # @return [Integer]
    attr_reader :rotY
    
    # depth or chart as % of chart width
    # must be between 20% and 2000%
    # @return [String]
    attr_reader :depthPercent
    
    # Chart axis are at right angles
    # @return [Boolean]
    attr_reader :rAngAx
    
    # field of view angle
    # @return [Integer]
    attr_reader :perspective

    # Creates a new View3D for charts
    # @option options [Integer] rotX
    # @option options [String] hPercent
    # @option options [Integer] rotY
    # @option options [String] depthPercent
    # @option options [Boolean] rAngAx
    # @option options [Integer] perspective
    def initialize(options={})
      @rotX, @hPercent, @rotY, @depthPercent, @rAngAx, @perspective  = nil, nil, nil, nil, nil, nil 
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end    
    end

    # @see rotX
    def rotX=(v) DataTypeValidator.validate "#{self.class}.rotX", [Integer, Fixnum], v, lambda {|arg| arg >= -90 && arg <= 90 }; @rotX = v; end

    # @see hPercent
    def hPercent=(v) RegexValidator.validate "#{self.class}.rotX", H_PERCENT_REGEX, v; @hPercent = v; end

    # @see rotY
    def rotY=(v) DataTypeValidator.validate "#{self.class}.rotY", [Integer, Fixnum], v, lambda {|arg| arg >= 0 && arg <= 360 }; @rotY = v; end

    # @see depthPercent
    def depthPercent=(v) RegexValidator.validate "#{self.class}.depthPercent", DEPTH_PERCENT_REGEX, v; @depthPercent = v; end

    # @see rAngAx
    def rAngAx=(v) Axlsx::validate_boolean(v); @rAngAx = v; end

    # @see perspective
    def perspective=(v) DataTypeValidator.validate "#{self.class}.perspective", [Integer, Fixnum], v, lambda {|arg| arg >= 0 && arg <= 240 }; @perspective = v; end

    # Serializes the view3D properties
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml[:c].view3D {
        xml[:c].rotX :val=>@rotX unless @rotX.nil?
        xml[:c].hPercent :val=>@hPercent unless @hPercent.nil?
        xml[:c].rotY :val=>@rotY unless @rotY.nil?
        xml[:c].depthPercent :val=>@depthPercent unless @depthPercent.nil?
        xml[:c].rAngAx :val=>@rAngAx unless @rAngAx.nil?
        xml[:c].perspective :val=>@perspective unless @perspective.nil?
      }
    end
  end
end
