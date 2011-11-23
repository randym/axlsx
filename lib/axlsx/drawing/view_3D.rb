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
    attr_accessor :rotX
    
    # height of chart as % of chart
    # must be between 5% and 500%
    # @return [String]
    attr_accessor :hPercent
    
    # y rotation for the chart
    # must be between 0 and 360
    # @return [Integer]
    attr_accessor :rotY
    
    # depth or chart as % of chart width
    # must be between 20% and 2000%
    # @return [String]
    attr_accessor :depthPercent
    
    # Chart axis are at right angles
    # @return [Boolean]
    attr_accessor :rAngAx
    
    # field of view angle
    # @return [Integer]
    attr_accessor :perspective

    # Creates a new View3D for charts
    # @option options [Integer] rotX
    # @option options [String] hPercent
    # @option options [Integer] rotY
    # @option options [String] depthPercent
    # @option options [Boolean] rAngAx
    # @option options [Integer] perspective
    def initialize(options={})
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end    
    end

    def rotX=(v) DataTypeValidator.validate "#{self.class}.rotX", [Integer, Fixnum], v, lambda {|v| v >= -90 && v <= 90 }; @rotX = v; end

    def hPercent=(v) RegexValidator.validate "#{self.class}.rotX", H_PERCENT_REGEX, v; @hPercent = v; end

    def rotY=(v) DataTypeValidator.validate "#{self.class}.rotY", [Integer, Fixnum], v, lambda {|v| v >= 0 && v <= 360 }; @rotY = v; end

    def depthPercent=(v) RegexValidator.validate "#{self.class}.depthPercent", DEPTH_PERCENT_REGEX, v; @depthPercent = v; end

    def rAngAx=(v) Axlsx::validate_boolean(v); @rAngAx = v; end

    def perspective=(v) DataTypeValidator.validate "#{self.class}.perspective", [Integer, Fixnum], v, lambda {|v| v >= 0 && v <= 240 }; @perspective = v; end

    # Serializes the view3D properties
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml.send('c:view3D') {
        xml.send('c:rotX', :val=>@rotX) unless @rotX.nil?
        xml.send('c:hPercent', :val=>@hPercent) unless @hPercent.nil?
        xml.send('c:rotY', :val=>@rotY) unless @rotY.nil?
        xml.send('c:depthPercent', :val=>@depthPercent) unless @depthPercent.nil?
        xml.send('c:rAngAx', :val=>@rAngAx) unless @rAngAx.nil?
        xml.send('c:perspective', :val=>@perspective) unless @perspective.nil?
      }
    end
  end
end
