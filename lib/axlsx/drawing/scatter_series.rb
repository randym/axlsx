# encoding: UTF-8
module Axlsx

  # A ScatterSeries defines the x and y position of data in the chart
  # @note The recommended way to manage series is to use Chart#add_series
  # @see Worksheet#add_chart
  # @see Chart#add_series
  # @see examples/example.rb
  class ScatterSeries < Series

    # The x data for this series.
    # @return [NamedAxisData]
    attr_reader :xData

    # The y data for this series.
    # @return [NamedAxisData]
    attr_reader :yData

    # The fill color for this series.
    # Red, green, and blue is expressed as sequence of hex digits, RRGGBB. A perceptual gamma of 2.2 is used.
    # @return [String]
    attr_reader :color

    # @return [String]
    attr_reader :ln_width

    # Line smoothing between data points
    # @return [Boolean]
    attr_reader :smooth

    # Creates a new ScatterSeries
    def initialize(chart, options={})
      @xData, @yData = nil
      if options[:smooth].nil?
        # If caller hasn't specified smoothing or not, turn smoothing on or off based on scatter style
        @smooth = [:smooth, :smoothMarker].include?(chart.scatter_style)
      else
        # Set smoothing according to the option provided
        Axlsx::validate_boolean(options[:smooth])
        @smooth = options[:smooth]
      end
      @ln_width = options[:ln_width] unless options[:ln_width].nil?
      super(chart, options)
      @xData = AxDataSource.new(:tag_name => :xVal, :data => options[:xData]) unless options[:xData].nil?
      @yData = NumDataSource.new({:tag_name => :yVal, :data => options[:yData]}) unless options[:yData].nil?
    end

    # @see color
    def color=(v)
      @color = v
    end

    # @see smooth
    def smooth=(v)
      Axlsx::validate_boolean(v)
      @smooth = v
    end

    # @see ln_width
    def ln_width=(v)
      @ln_width = v
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      super(str) do
        # needs to override the super color here to push in ln/and something else!
        if color
          str << '<c:spPr><a:solidFill>'
          str << ('<a:srgbClr val="' << color << '"/>')
          str << '</a:solidFill>'
          str << '<a:ln><a:solidFill>'
          str << ('<a:srgbClr val="' << color << '"/></a:solidFill></a:ln>')
          str << '</c:spPr>'
          str << '<c:marker>'
          str << '<c:spPr><a:solidFill>'
          str << ('<a:srgbClr val="' << color << '"/>')
          str << '</a:solidFill>'
          str << '<a:ln><a:solidFill>'
          str << ('<a:srgbClr val="' << color << '"/></a:solidFill></a:ln>')
          str << '</c:spPr>'
          str << '</c:marker>'
        end
        if ln_width
          str << '<c:spPr>'
          str << '<a:ln w="' << ln_width.to_s << '"/>'
          str << '</c:spPr>'
        end
        @xData.to_xml_string(str) unless @xData.nil?
        @yData.to_xml_string(str) unless @yData.nil?
        str << ('<c:smooth val="' << ((smooth) ? '1' : '0') << '"/>')
      end
      str
    end
  end
end
