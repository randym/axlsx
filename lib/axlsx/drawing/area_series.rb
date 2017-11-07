# encoding: UTF-8
module Axlsx
  # A AreaSeries defines the title, data and labels for line charts
  # @note The recommended way to manage series is to use Chart#add_series
  # @see Worksheet#add_chart
  # @see Chart#add_series
  class AreaSeries < Series

    # The data for this series.
    # @return [ValAxisData]
    attr_reader :data

    # The labels for this series.
    # @return [CatAxisData]
    attr_reader :labels

    # The fill color for this series.
    # Red, green, and blue is expressed as sequence of hex digits, RRGGBB. A perceptual gamma of 2.2 is used.
    # @return [String]
    attr_reader :color

    # show markers on values
    # @return [Boolean]
    attr_reader :show_marker

    # custom marker symbol
    # @return [String]
    attr_reader :marker_symbol

    # line smoothing on values
    # @return [Boolean]
    attr_reader :smooth

    # Creates a new series
    # @option options [Array, SimpleTypedList] data
    # @option options [Array, SimpleTypedList] labels
    # @param [Chart] chart
    def initialize(chart, options={})
      @show_marker = false
      @marker_symbol = options[:marker_symbol] ? options[:marker_symbol] : :default
      @smooth = false
      @labels, @data = nil, nil
      super(chart, options)
      @labels = AxDataSource.new(:data => options[:labels]) unless options[:labels].nil?
      @data = NumDataSource.new(options) unless options[:data].nil?
    end

    # @see color
    def color=(v)
      @color = v
    end

    # @see show_marker
    def show_marker=(v)
      Axlsx::validate_boolean(v)
      @show_marker = v
    end

    # @see marker_symbol
    def marker_symbol=(v)
      Axlsx::validate_marker_symbol(v)
      @marker_symbol = v
    end

    # @see smooth
    def smooth=(v)
      Axlsx::validate_boolean(v)
      @smooth = v
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      super(str) do
        if color
          str << '<c:spPr><a:solidFill>'
          str << ('<a:srgbClr val="' << color << '"/>')
          str << '</a:solidFill>'
          str << '<a:ln w="28800">'
          str << '<a:solidFill>'
          str << ('<a:srgbClr val="' << color << '"/>')
          str << '</a:solidFill>'
          str << '</a:ln>'
          str << '<a:round/>'
          str << '</c:spPr>'
        end

        if !@show_marker
          str << '<c:marker><c:symbol val="none"/></c:marker>'
        elsif @marker_symbol != :default
          str << '<c:marker><c:symbol val="' + @marker_symbol.to_s + '"/></c:marker>'
        end

        @labels.to_xml_string(str) unless @labels.nil?
        @data.to_xml_string(str) unless @data.nil?
        str << ('<c:smooth val="' << ((smooth) ? '1' : '0') << '"/>')
      end
    end

    private

    # assigns the data for this series
    def data=(v) DataTypeValidator.validate "Series.data", [NumDataSource], v; @data = v; end

    # assigns the labels for this series
    def labels=(v) DataTypeValidator.validate "Series.labels", [AxDataSource], v; @labels = v; end

  end
end
