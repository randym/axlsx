module Axlsx
  # A Title stores information about the title of a chart
  class Title
    
    # The text to be shown. Setting this property directly with a string will remove the cell reference.
    # @return [String]
    attr_accessor :text

    # The cell that holds the text for the title. Setting this property will automatically update the text attribute.
    # @return [Cell]
    attr_accessor :cell

    # Creates a new Title object
    # @param [String, Cell] title The cell or string to be used for the chart's title
    def initialize(title="")
      self.cell = title if title.is_a?(Cell)
      self.text = title.to_s unless title.is_a?(Cell)
    end
    
    def text=(v) 
      DataTypeValidator.validate 'Title.text', String, v
      @text = v
      @cell = nil
      v
    end

    def cell=(v)
      DataTypeValidator.validate 'Title.text', Cell, v
      @cell = v
      @text = v.value.to_s      
      v
    end

    # Not implemented at this time.
    #def layout=(v) DataTypeValidator.validate 'Title.layout', Layout, v; @layout = v; end
    #def overlay=(v) Axlsx::validate_boolean v; @overlay=v; end
    #def spPr=(v) DataTypeValidator.validate 'Title.spPr', SpPr, v; @spPr = v; end
    
    # Serializes the chart title
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml.send('c:title') {
        xml.send('c:tx') {
          xml.send('c:strRef') {
            xml.send('c:f', range)
            xml.send('c:strCache') {
              xml.send('c:ptCount', :val=>1)
              xml.send('c:pt', :idx=>0) {
                xml.send('c:v', @text)
              }
            }
          }
        }
      }      
    end
    
    private 

    # returns the excel style abslute reference for the title when title is a Cell object
    # @return [String]
    def range
      return "" unless @data.is_a?(Cell)
      "#{@data.row.worksheet.name}!#{data.row.r_abs}"
    end

  end
end
