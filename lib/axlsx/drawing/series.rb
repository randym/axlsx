module Axlsx
  # A Series defines the title, data and labels for chart data.
  # @note The recommended way to manage series is to use Chart#add_series
  # @see Worksheet#add_chart
  # @see Chart#add_series
  class Series

    # The chart that owns this series
    # @return [Chart]
    attr_reader :chart

    # The index of this series in the chart's series.
    # @return [Integer]
    attr_reader :index

    # The order of this series in the chart's series.
    # @return [Integer]
    attr_accessor :order

    # The title of the series
    # @return [SeriesTitle]
    attr_accessor :title

    # Creates a new series
    # @param [Chart] chart
    # @option options [Integer] order
    # @option options [String] title
    def initialize(chart, options={})
      self.chart = chart
      @chart.series << self
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
    end    
    
    # retrieves the series index in the chart's series collection
    def index
      @chart.series.index(self)
    end

    def order=(v)  Axlsx::validate_unsigned_int(v); @order = v; end

    def order
      @order || index
    end

    def title=(v) 
      v = SeriesTitle.new(v) if v.is_a?(String) || v.is_a?(Cell)
      DataTypeValidator.validate "#{self.class}.title", SeriesTitle, v
      @title = v
    end
   
    private 
    
    # assigns the chart for this series
    def chart=(v)  DataTypeValidator.validate "Series.chart", Chart, v; @chart = v; end    

    # Serializes the series
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml.send('c:ser') {
        xml.send('c:idx', :val=>index)
        xml.send('c:order', :val=>order || index)        
        title.to_xml(xml) unless title.nil?
        yield xml if block_given?
      }
    end

  end

end
