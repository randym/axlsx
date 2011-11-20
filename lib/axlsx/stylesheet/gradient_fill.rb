# -*- coding: utf-8 -*-
module Axlsx
  # A GradientFill defines the color and positioning for gradiant cell fill.
  # @see Open Office XML Part 1 §18.8.24  
  class GradientFill

    # The type of gradient.
    # @note 
    #  valid options are
    #   :linear
    #   :path
    # @return [Symbol]
    attr_accessor :type

    # Angle of the linear gradient
    # @return [Float]
    attr_accessor :degree

    # Percentage format left
    # @return [Float]
    attr_accessor :left

    # Percentage format right
    # @return [Float]
    attr_accessor :right

    # Percentage format top
    # @return [Float]
    attr_accessor :top 

    # Percentage format bottom
    # @return [Float]
    attr_accessor :bottom

    # Collection of stop objects
    # @return [SimpleTypedList]
    attr_reader :stop
    
    # Creates a new GradientFill object
    # @option options [Symbol] type
    # @option options [Float] degree
    # @option options [Float] left
    # @option options [Float] right
    # @option options [Float] top
    # @option options [Float] bottom
    def initialize(options={})
      options[:type] ||= :linear
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? o[0]
      end
      @stop = SimpleTypedList.new GradientStop
    end

    def type=(v) Axlsx::validate_gradient_type v; @type = v end    
    def degree=(v) Axlsx::validate_float v; @degree = v end    
    def left=(v) DataTypeValidator.validate "GradientFill.left", Float, v, lambda { |v| v >= 0.0 && v <= 1.0}; @left = v end    
    def right=(v) DataTypeValidator.validate "GradientFill.right", Float, v, lambda { |v| v >= 0.0 && v <= 1.0}; @right = v end    
    def top=(v) DataTypeValidator.validate "GradientFill.top", Float, v, lambda { |v| v >= 0.0 && v <= 1.0}; @top = v end    
    def bottom=(v) DataTypeValidator.validate "GradientFill.bottom", Float, v, lambda { |v| v >= 0.0 && v <= 1.0}; @bottom= v end    

    # Serializes the gradientFill
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml.gradientFill(self.instance_values.reject { |k,v| k.to_sym == :stop }) {
        @stop.each { |s| s.to_xml(xml) }
      }
    end
  end
end
