# <xsd:complexType name="CT_BookView">
#     <xsd:sequence>
#       <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
#     </xsd:sequence>
#     <xsd:attribute name="visibility" type="ST_Visibility" use="optional" default="visible"/>
#     <xsd:attribute name="minimized" type="xsd:boolean" use="optional" default="false"/>
#     <xsd:attribute name="showHorizontalScroll" type="xsd:boolean" use="optional" default="true"/>
#     <xsd:attribute name="showVerticalScroll" type="xsd:boolean" use="optional" default="true"/>
#     <xsd:attribute name="showSheetTabs" type="xsd:boolean" use="optional" default="true"/>
#     <xsd:attribute name="xWindow" type="xsd:int" use="optional"/>
#     <xsd:attribute name="yWindow" type="xsd:int" use="optional"/>
#     <xsd:attribute name="windowWidth" type="xsd:unsignedInt" use="optional"/>
#     <xsd:attribute name="windowHeight" type="xsd:unsignedInt" use="optional"/>
#     <xsd:attribute name="tabRatio" type="xsd:unsignedInt" use="optional" default="600"/>
#     <xsd:attribute name="firstSheet" type="xsd:unsignedInt" use="optional" default="0"/>
#     <xsd:attribute name="activeTab" type="xsd:unsignedInt" use="optional" default="0"/>
#     <xsd:attribute name="autoFilterDateGrouping" type="xsd:boolean" use="optional"
#     default="true"/>
# </xsd:complexType>

module Axlsx

  # A BookView defines the display properties for a workbook.
  # Units for window widths and other dimensions are expressed in twips.
  # Twip measurements are portable between different display resolutions.
  # The formula is (screen pixels) * (20 * 72) / (logical device dpi),
  # where the logical device dpi can be different for x and y coordinates.
  class WorkbookView

    include Axlsx::SerializedAttributes
    include Axlsx::OptionsParser
    include Axlsx::Accessors


    # Creates a new BookView object
    # @param [Hash] options  A hash of key/value pairs that will be mapped to this instances attributes.
    # @option [Symbol] visibility Specifies visible state of the workbook window. The default value for this attribute is :visible.
    # @option [Boolean] minimized Specifies a boolean value that indicates whether the workbook window is minimized.
    # @option [Boolean] show_horizontal_scroll Specifies a boolean value that indicates whether to display the horizontal scroll bar in the user interface.
    # @option [Boolean] show_vertical_scroll Specifies a boolean value that indicates whether to display the vertical scroll bar.
    # @option [Boolean] show_sheet_tabs Specifies a boolean value that indicates whether to display the sheet tabs in the user interface.
    # @option [Integer] tab_ratio Specifies ratio between the workbook tabs bar and the horizontal scroll bar.
    # @option [Integer] first_sheet Specifies the index to the first sheet in this book view.
    # @option [Integer] active_tab Specifies an unsignedInt that contains the index to the active sheet in this book view.
    # @option [Integer] x_window Specifies the X coordinate for the upper left corner of the workbook window. The unit of measurement for this value is twips.
    # @option [Integer] y_window Specifies the Y coordinate for the upper left corner of the workbook window. The unit of measurement for this value is twips.
    # @option [Integer] window_width Specifies the width of the workbook window. The unit of measurement for this value is twips.
    # @option [Integer] window_height Specifies the height of the workbook window. The unit of measurement for this value is twips.
    # @option [Boolean] auto_filter_date_grouping Specifies a boolean value that indicates whether to group dates when presenting the user with filtering options in the user interface.
    def initialize(options={})
      parse_options options
      yield self if block_given?
    end


    unsigned_int_attr_accessor :x_window, :y_window, :window_width, :window_height,
                               :tab_ratio, :first_sheet, :active_tab

    validated_attr_accessor  [:visibility], :validate_view_visibility

    serializable_attributes :visibility, :minimized,
                            :show_horizontal_scroll, :show_vertical_scroll,
                            :show_sheet_tabs, :tab_ratio, :first_sheet, :active_tab,
                            :x_window, :y_window, :window_width, :window_height,
                            :auto_filter_date_grouping

    boolean_attr_accessor :minimized, :show_horizontal_scroll, :show_vertical_scroll,
                          :show_sheet_tabs, :auto_filter_date_grouping


    # Serialize the WorkbookView
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
       str << '<workbookView '
       serialized_attributes str
       str << '></workbookView>'
    end
  end
end
