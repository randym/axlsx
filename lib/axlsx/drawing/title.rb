# encoding: UTF-8
# frozen_string_literal: true
module Axlsx
  # A Title stores information about the title of a chart
  class Title

    # The text to be shown. Setting this property directly with a string will remove the cell reference.
    # @return [String]
    attr_reader :text

    # Text size property
    # @return [String]
    attr_reader :text_size

    # The cell that holds the text for the title. Setting this property will automatically update the text attribute.
    # @return [Cell]
    attr_reader :cell

    # Creates a new Title object
    # @param [String, Cell] title The cell or string to be used for the chart's title
    def initialize(title="", title_size="")
      self.cell = title if title.is_a?(Cell)
      self.text = title.to_s unless title.is_a?(Cell)
      if title_size.to_s.empty?
        self.text_size = "1600"
      else
        self.text_size = title_size.to_s
      end
    end

    # @see text
    def text=(v)
      DataTypeValidator.validate 'Title.text', String, v
      @text = v
      @cell = nil
      v
    end

    # @see text_size
    def text_size=(v)
      DataTypeValidator.validate 'Title.text_size', String, v
      @text_size = v
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

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = String.new)
      str << '<c:title>'
      unless @text.empty?
        str << '<c:tx>'
        if @cell.is_a?(Cell)
          str << '<c:strRef>'\
                 "<c:f>#{Axlsx::cell_range([@cell])}</c:f>"\
                 '<c:strCache>'\
                 '<c:ptCount val="1"/>'\
                 '<c:pt idx="0">'\
                 "<c:v>#{@text}</c:v>"\
                 '</c:pt>'\
                 '</c:strCache>'\
                 '</c:strRef>'
        else
          str << '<c:rich>'\
                   '<a:bodyPr/>'\
                   '<a:lstStyle/>'\
                   '<a:p>'\
                     '<a:r>'\
                       "<a:rPr sz=\"#{@text_size}\"/>"\
                       "<a:t>#{@text}</a:t>"\
                     '</a:r>'\
                   '</a:p>'\
                 '</c:rich>'
        end
        str << '</c:tx>'
      end
      str << '<c:layout/>'\
             '<c:overlay val="0"/>'\
             '</c:title>'
    end

  end
end
