module Axlsx
  # A BarSeries defines the title, data and labels for bar charts
  # @note The recommended way to manage series is to use Chart#add_series
  # @see Worksheet#add_chart
  # @see Chart#add_series
  class BarSeries < Series

  
    # The data for this series. 
    # @return [Array, SimpleTypedList]
    attr_reader :data

    # The labels for this series.
    # @return [Array, SimpleTypedList]
    attr_reader :labels

    # The shabe of the bars or columns
    # must be one of  [:percentStacked, :clustered, :standard, :stacked]
    # @return [Symbol]
    attr_accessor :shape

    # Creates a new series
    # @option options [Array, SimpleTypedList] data
    # @option options [Array, SimpleTypedList] labels
    # @option options [String] title
    # @option options [String] shape
    # @param [Chart] chart
    def initialize(chart, options={})
      @shape = :box
      super(chart, options)
      self.data = options[:data]  || []
      self.labels = options[:labels] || []
    end 

    def shape=(v) 
      RestrictionValidator.validate "BarSeries.shape", [:cone, :coneToMax, :box, :cylinder, :pyramid, :pyramidToMax], v
      @shape = v
    end

    # Serializes the series
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      super(xml) do |xml|
        if !labels.empty?
          xml.send('c:cat') {
            xml.send('c:strRef') {
              xml.send('c:f', Axlsx::cell_range(labels))
              xml.send('c:strCache') {
                xml.send('c:ptCount', :val=>labels.size)
                labels.each_with_index do |cell, index|
                  v = cell.is_a?(Cell) ? cell.value : cell
                  xml.send('c:pt', :idx=>index) {
                    xml.send('c:v', v)
                  }                          
                end
              }
            }
          }
        end
        xml.send('c:val') {
          xml.send('c:numRef') {
            xml.send('c:f', Axlsx::cell_range(data))
            xml.send('c:numCache') {
              xml.send('c:formatCode', 'General')
              xml.send('c:ptCount', :val=>data.size)
              data.each_with_index do |cell, index|
                v = cell.is_a?(Cell) ? cell.value : cell
                xml.send('c:pt', :idx=>index) {
                  xml.send('c:v', v) 
                }
              end
            }                        
          }
        }
      end      
    end

    
    private 


    # assigns the data for this series
    def data=(v) DataTypeValidator.validate "Series.data", [Array, SimpleTypedList], v; @data = v; end

    # assigns the labels for this series
    def labels=(v) DataTypeValidator.validate "Series.labels", [Array, SimpleTypedList], v; @labels = v; end

  end

end
