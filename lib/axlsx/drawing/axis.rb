# encoding: UTF-8
module Axlsx

  # the access class defines common properties and values for a chart axis.
  class Axis

    include Axlsx::OptionsParser

    # Creates an Axis object
    # @option options [Axis] cross_axis the perpendicular axis
    # @option options [Symbol] ax_pos
    # @option options [Symbol] crosses
    # @option options [Symbol] tick_lbl_pos
    # @raise [ArgumentError] If axi_id or cross_ax are not unsigned integers
    def initialize(options={})
      @id = rand(8 ** 8)
      @format_code = "General"
      @delete = @label_rotation = 0
      @scaling = Scaling.new(:orientation=>:minMax)
      @title = @color = nil
      self.ax_pos = :b
      self.tick_lbl_pos = :nextTo
      self.format_code = "General"
      self.crosses = :autoZero
      self.gridlines = true
      parse_options options
    end

    # the fill color to use in the axis shape properties. This should be a 6 character long hex string
    # e.g. FF0000 for red
    # @return [String]
    attr_reader :color

    # the id of the axis.
    # @return [Integer]
    attr_reader :id
    alias :axID :id

    # The perpendicular axis
    # @return [Integer]
    attr_reader :cross_axis
    alias :crossAx :cross_axis

    # The scaling of the axis
    # @see Scaling
    # @return [Scaling]
    attr_reader :scaling

    # The position of the axis
    # must be one of [:l, :r, :t, :b]
    # @return [Symbol]
    attr_reader :ax_pos
    alias :axPos :ax_pos

    # the position of the tick labels
    # must be one of [:nextTo, :high, :low]
    # @return [Symbol]
    attr_reader :tick_lbl_pos
    alias :tickLblPos :tick_lbl_pos

    # The number format format code for this axis
    # default :General
    # @return [String]
    attr_reader :format_code

    # specifies how the perpendicular axis is crossed
    # must be one of [:autoZero, :min, :max]
    # @return [Symbol]
    attr_reader :crosses

    # specifies how the degree of label rotation
    # @return [Integer]
    attr_reader :label_rotation

    # specifies if gridlines should be shown in the chart
    # @return [Boolean]
    attr_reader :gridlines

    # specifies if gridlines should be shown in the chart
    # @return [Boolean]
    attr_reader :delete

    # the title for the axis. This can be a cell or a fixed string.
    attr_reader :title

    # The color for this axis. This value is used when rendering the axis line in the chart. 
    # colors should be in 6 character rbg format
    # @return [String] the rbg color assinged.
    # @see color
    def color=(color_rgb)
      @color = color_rgb
    end
    
    # The crossing axis for this axis
    # @param [Axis] axis
    def cross_axis=(axis)
       DataTypeValidator.validate "#{self.class}.cross_axis", [Axis], axis
       @cross_axis = axis
    end

    # The position of the axis
    # must be one of [:l, :r, :t, :b]
    def ax_pos=(v) RestrictionValidator.validate "#{self.class}.ax_pos", [:l, :r, :b, :t], v; @ax_pos = v; end
    alias :axPos= :ax_pos=

    # the position of the tick labels
    # must be one of [:nextTo, :high, :low1]
    def tick_lbl_pos=(v) RestrictionValidator.validate "#{self.class}.tick_lbl_pos", [:nextTo, :high, :low, :none], v; @tick_lbl_pos = v; end
    alias :tickLblPos= :tick_lbl_pos=

    # The number format format code for this axis
    # default :General
    def format_code=(v) Axlsx::validate_string(v); @format_code = v; end

    # Specify if gridlines should be shown for this axis
    # default true
    def gridlines=(v) Axlsx::validate_boolean(v); @gridlines = v; end

    # Specify if axis should be removed from the chart
    # default false
    def delete=(v) Axlsx::validate_boolean(v); @delete = v; end

    # specifies how the perpendicular axis is crossed
    # must be one of [:autoZero, :min, :max]
    def crosses=(v) RestrictionValidator.validate "#{self.class}.crosses", [:autoZero, :min, :max], v; @crosses = v; end

    # Specify the degree of label rotation to apply to labels
    # default true
    def label_rotation=(v)
      Axlsx::validate_int(v)
      adjusted = v.to_i * 60000
      Axlsx::validate_angle(adjusted)
      @label_rotation = adjusted
    end

    # The title object for the chart.
    # @param [String, Cell] v
    # @return [Title]
    def title=(v)
      DataTypeValidator.validate "#{self.class}.title", [String, Cell], v
      @title ||= Title.new
      if v.is_a?(String)
        @title.text = v
      elsif v.is_a?(Cell)
        @title.cell = v
      end
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<c:axId val="' << @id.to_s << '"/>'
      @scaling.to_xml_string str
      str << '<c:delete val="'<< @delete.to_s << '"/>'
      str << '<c:axPos val="' << @ax_pos.to_s << '"/>'
      str << '<c:majorGridlines>'
      # TODO shape properties need to be extracted into a class
      if gridlines == false
        str << '<c:spPr>'
        str << '<a:ln>'
        str << '<a:noFill/>'
        str << '</a:ln>'
        str << '</c:spPr>'
      end
      str << '</c:majorGridlines>'
      @title.to_xml_string(str) unless @title == nil
      str << '<c:numFmt formatCode="' << @format_code << '" sourceLinked="1"/>'
      str << '<c:majorTickMark val="none"/>'
      str << '<c:minorTickMark val="none"/>'
      str << '<c:tickLblPos val="' << @tick_lbl_pos.to_s << '"/>'
      # TODO - this is also being used for series colors
      # time to extract this into a class spPr - Shape Properties
      if @color
        str << '<c:spPr><a:ln><a:solidFill>'
        str << '<a:srgbClr val="' << @color << '"/>'
        str << '</a:solidFill></a:ln></c:spPr>'
      end
      # some potential value in implementing this in full. Very detailed!
      str << '<c:txPr><a:bodyPr rot="' << @label_rotation.to_s << '"/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr/></a:p></c:txPr>'
      str << '<c:crossAx val="' << @cross_axis.id.to_s << '"/>'
      str << '<c:crosses val="' << @crosses.to_s << '"/>'
    end

  end

end
