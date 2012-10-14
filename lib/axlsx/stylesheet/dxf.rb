# encoding: UTF-8
module Axlsx
  # The Dxf class defines an incremental formatting record for use in Styles. The recommended way to manage styles for your workbook is with Styles#add_style
  # @see Styles#add_style
  class Dxf

    include Axlsx::OptionsParser

    # The order in which the child elements is put in the XML seems to
    # be important for Excel
    CHILD_ELEMENTS = [:font, :numFmt, :fill, :alignment, :border, :protection]
    #does not support extList (ExtensionList)

    # The cell alignment for this style
    # @return [CellAlignment]
    # @see CellAlignment
    attr_reader :alignment

    # The cell protection for this style
    # @return [CellProtection]
    # @see CellProtection
    attr_reader :protection

    # the child NumFmt to be used to this style
    # @return [NumFmt]
    attr_reader :numFmt

    # the child font to be used for this style
    # @return [Font]
    attr_reader :font

    # the child fill to be used in this style
    # @return [Fill]
    attr_reader :fill

    # the border to be used in this style
    # @return [Border]
    attr_reader :border

    # Creates a new Xf object
    # @option options [Border] border
    # @option options [NumFmt] numFmt
    # @option options [Fill] fill
    # @option options [Font] font
    # @option options [CellAlignment] alignment
    # @option options [CellProtection] protection
    def initialize(options={})
      parse_options options
    end

    # @see Dxf#alignment
    def alignment=(v) DataTypeValidator.validate "Dxf.alignment", CellAlignment, v; @alignment = v end
    # @see protection
    def protection=(v) DataTypeValidator.validate "Dxf.protection", CellProtection, v; @protection = v end
    # @see numFmt
    def numFmt=(v) DataTypeValidator.validate "Dxf.numFmt", NumFmt, v; @numFmt = v end
    # @see font
    def font=(v) DataTypeValidator.validate "Dxf.font", Font, v; @font = v end
    # @see border
    def border=(v) DataTypeValidator.validate "Dxf.border", Border, v; @border = v end
    # @see fill
    def fill=(v) DataTypeValidator.validate "Dxf.fill", Fill, v; @fill = v end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<dxf>'
      # Dxf elements have no attributes. All of the instance variables
      # are child elements.
      CHILD_ELEMENTS.each do |element|
        self.send(element).to_xml_string(str) if self.send(element)
      end
      str << '</dxf>'
    end

  end

end
