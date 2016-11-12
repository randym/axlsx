# encoding: UTF-8
module Axlsx
  # 3D attributes for a chart.
  class View3D

    include Axlsx::OptionsParser

    # Creates a new View3D for charts
    # @option options [Integer] rot_x
    # @option options [String] h_percent
    # @option options [Integer] rot_y
    # @option options [String] depth_percent
    # @option options [Boolean] r_ang_ax
    # @option options [Integer] perspective
    def initialize(options={})
      @rot_x, @h_percent, @rot_y, @depth_percent, @r_ang_ax, @perspective  = nil, nil, nil, nil, nil, nil
      parse_options options
    end

    # Validation for hPercent
    H_PERCENT_REGEX = /0*(([5-9])|([1-9][0-9])|([1-4][0-9][0-9])|500)/

    # validation for depthPercent
    DEPTH_PERCENT_REGEX = /0*(([2-9][0-9])|([1-9][0-9][0-9])|(1[0-9][0-9][0-9])|2000)/

    # x rotation for the chart
    # must be between -90 and 90
    # @return [Integer]
    attr_reader :rot_x
    alias :rotX :rot_x

    # height of chart as % of chart width
    # must be between 5% and 500%
    # @return [String]
    attr_reader :h_percent
    alias :hPercent :h_percent

    # y rotation for the chart
    # must be between 0 and 360
    # @return [Integer]
    attr_reader :rot_y
    alias :rotY :rot_y

    # depth or chart as % of chart width
    # must be between 20% and 2000%
    # @return [String]
    attr_reader :depth_percent
    alias :depthPercent :depth_percent

    # Chart axis are at right angles
    # @return [Boolean]
    attr_reader :r_ang_ax
    alias :rAngAx :r_ang_ax

    # field of view angle
    # @return [Integer]
    attr_reader :perspective

    # @see rot_x
    def rot_x=(v)
      RangeValidator.validate "View3D.rot_x", -90, 90, v
      @rot_x = v
    end
    alias :rotX= :rot_x=

      # @see h_percent
      def h_percent=(v)
        RegexValidator.validate "#{self.class}.h_percent", H_PERCENT_REGEX, v
        @h_percent = v
      end
    alias :hPercent= :h_percent=

      # @see rot_y
      def rot_y=(v) 
        RangeValidator.validate "View3D.rot_y", 0, 360, v
        @rot_y = v
      end
    alias :rotY= :rot_y=

      # @see depth_percent
      def depth_percent=(v) RegexValidator.validate "#{self.class}.depth_percent", DEPTH_PERCENT_REGEX, v; @depth_percent = v; end
    alias :depthPercent= :depth_percent=

      # @see r_ang_ax
      def r_ang_ax=(v) Axlsx::validate_boolean(v); @r_ang_ax = v; end
    alias :rAngAx= :r_ang_ax=

      # @see perspective
      def perspective=(v)
        RangeValidator.validate "View3D.perspective", 0, 240, v
        @perspective = v
      end

    # DataTypeValidator.validate "#{self.class}.perspective", [Integer], v, lambda {|arg| arg >= 0 && arg <= 240 }; @perspective = v; end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<c:view3D>'
      %w(rot_x h_percent rot_y depth_percent r_ang_ax perspective).each do |key|
        str << element_for_attribute(key, 'c')
      end
      str << '</c:view3D>'
    end

    private
    # Note: move this to Axlsx module if we find the smae pattern elsewhere.
    def element_for_attribute(name, namespace='')
      val = instance_values[name]
      return "" if val == nil
      "<%s:%s val='%s'/>" % [namespace, Axlsx::camel(name, false), val]
    end
  end
end
