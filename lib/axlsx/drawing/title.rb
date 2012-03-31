# encoding: UTF-8
module Axlsx
  # A Title stores information about the title of a chart
  class Title

    # The text to be shown. Setting this property directly with a string will remove the cell reference.
    # @return [String]
    attr_reader :text

    # The cell that holds the text for the title. Setting this property will automatically update the text attribute.
    # @return [Cell]
    attr_reader :cell

    # Creates a new Title object
    # @param [String, Cell] title The cell or string to be used for the chart's title
    def initialize(title="")
      self.cell = title if title.is_a?(Cell)
      self.text = title.to_s unless title.is_a?(Cell)
    end

    # @see text
    def text=(v)
      DataTypeValidator.validate 'Title.text', String, v
      @text = v
      @cell = nil
      v
    end

    # @see cell
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

    def to_xml_string(str = '')
      str << '<c:title>'
      unless @text.empty?
        str << '<c:tx>'
        str << '<c:strRef>'
        str << '<c:f>' << Axlsx::cell_range([@cell]) << '</c:f>'
        str << '<c:strCache>'
        str << '<c:ptCount val="1"/>'
        str << '<c:pt idx="0">'
        str << '<c:v>' << @text << '</c:v>'
        str << '</c:pt>'
        str << '</c:strCache>'
        str << '</c:strRef>'
        str << '</c:tx>'
      end
      str << '</c:title>'
    end

    # Serializes the chart title
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml[:c].title {
        unless @text.empty?
          xml[:c].tx {
            xml[:c].strRef {
              xml[:c].f Axlsx::cell_range([@cell])
              xml[:c].strCache {
                xml[:c].ptCount :val=>1
                xml[:c].pt(:idx=>0) {
                  xml[:c].v @text
                }
              }
            }
          }
        end
        xml[:c].layout
        xml[:c].overlay :val=>0
      }
    end

  end
end
