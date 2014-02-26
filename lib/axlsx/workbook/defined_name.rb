#  <definedNames>
#    <definedName name="_xlnm.Print_Titles" localSheetId="0">Sheet1!$1:$1</definedName>
#  </definedNames>

#￼￼￼<xsd:complexType name="CT_DefinedName">
# <xsd:simpleContent>
#￼￼￼￼ <xsd:extension base="ST_Formula">
# <xsd:attribute name="name" type="s:ST_Xstring" use="required"/>
# <xsd:attribute name="comment" type="s:ST_Xstring" use="optional"/>
# <xsd:attribute name="customMenu" type="s:ST_Xstring" use="optional"/>
# <xsd:attribute name="description" type="s:ST_Xstring" use="optional"/>
# <xsd:attribute name="help" type="s:ST_Xstring" use="optional"/>
# <xsd:attribute name="statusBar" type="s:ST_Xstring" use="optional"/>
# <xsd:attribute name="localSheetId" type="xsd:unsignedInt" use="optional"/>
# <xsd:attribute name="hidden" type="xsd:boolean" use="optional" default="false"/>
# <xsd:attribute name="function" type="xsd:boolean" use="optional" default="false"/>
# <xsd:attribute name="vbProcedure" type="xsd:boolean" use="optional" default="false"/>
# <xsd:attribute name="xlm" type="xsd:boolean" use="optional" default="false"/>
# <xsd:attribute name="functionGroupId" type="xsd:unsignedInt" use="optional"/>
# <xsd:attribute name="shortcutKey" type="s:ST_Xstring" use="optional"/>
# <xsd:attribute name="publishToServer" type="xsd:boolean" use="optional"￼ default="false"/>
# <xsd:attribute name="workbookParameter" type="xsd:boolean" use="optional" default="false"/>
# </xsd:extenstion>
# </xsd:simpleContent>

module Axlsx
  # This element defines the defined names that are defined within this workbook.
  # Defined names are descriptive text that is used to represents a cell, range of cells, formula, or constant value.
  # Use easy-to-understand names, such as Products, to refer to hard to understand ranges, such as Sales!C20:C30.
  # A defined name in a formula can make it easier to understand the purpose of the formula.
  # @example
  #     The formula =SUM(FirstQuarterSales) might be easier to identify than =SUM(C20:C30
  #
  # Names are available to any sheet.
  # @example
  #     If the name ProjectedSales refers to the range A20:A30 on the first worksheet in a workbook,
  #     you can use the name ProjectedSales on any other sheet in the same workbook to refer to range A20:A30 on the first worksheet.
  # Names can also be used to represent formulas or values that do not change (constants).
  #
  # @example
  #     The name SalesTax can be used to represent the sales tax amount (such as 6.2 percent) applied to sales transactions.
  # You can also link to a defined name in another workbook, or define a name that refers to cells in another workbook.
  #
  # @example
  #     The formula =SUM(Sales.xls!ProjectedSales) refers to the named range ProjectedSales in the workbook named Sales.
  # A compliant producer or consumer considers a defined name in the range A1-XFD1048576 to be an error.
  # All other names outside this range can be defined as names and overrides a cell reference if an ambiguity exists.
  #
  # @example
  #     For clarification: LOG10 is always a cell reference, LOG10() is always formula, LOGO1000 can be a defined name that overrides a cell reference.
  class DefinedName
    include Axlsx::SerializedAttributes
    include Axlsx::OptionsParser
    include Axlsx::Accessors
    # creates a new DefinedName.
    # @param [String] formula - the formula the defined name references
    # @param [Hash] options - A hash of key/value pairs that will be mapped to this instances attributes.
    #
    # @option [String] name - Specifies the name that appears in the user interface for the defined name.
    #                         This attribute is required.
    #                         The following built-in names are defined in this SpreadsheetML specification:
    #                         Print
    #                           _xlnm.Print_Area: this defined name specifies the workbook's print area.
    #                           _xlnm.Print_Titles: this defined name specifies the row(s) or column(s) to repeat
    #                                                the top of each printed page.
    #                         Filter & Advanced Filter
    #                           _xlnm.Criteria: this defined name refers to a range containing the criteria values
    #                                  to be used in applying an advanced filter to a range of data.
    #                           _xlnm._FilterDatabase: can be one of the following
    #                                     a. this defined name refers to a range to which an advanced filter has been
    #                                         applied. This represents the source data range, unfiltered.
    #                                      b. This defined name refers to a range to which an AutoFilter has been
    #                                         applied.
    #                           _xlnm.Extract: this defined name refers to the range containing the filtered output
    #                                           values resulting from applying an advanced filter criteria to a source range.
    #                         Miscellaneous
    #                           _xlnm.Consolidate_Area: the defined name refers to a consolidation area.
    #                           _xlnm.Database: the range specified in the defined name is from a database data source.
    #                           _xlnm.Sheet_Title: the defined name refers to a sheet title.
    # @option [String] comment - A comment to optionally associate with the name
    # @option [String] custom_menu - The menu text for the defined name
    # @option [String] description - An optional description for the defined name
    # @option [String] help - The help topic to display for this defined name
    # @option [String] status_bar - The text to display on the application status bar when this defined name has focus
    # @option [String] local_sheet_id - Specifies the sheet index in this workbook where data from an external reference is displayed
    # @option [Boolean] hidden - Specifies a boolean value that indicates whether the defined name is hidden in the user interface.
    # @option [Boolean] function - Specifies a boolean value that indicates that the defined name refers to a user-defined function.
    #                             This attribute is used when there is an add-in or other code project associated with the file.
    # @option [Boolean] vb_proceedure - Specifies a boolean value that indicates whether the defined name is related to an external function, command, or other executable code.
    # @option [Boolean] xlm - Specifies a boolean value that indicates whether the defined name is related to an external function, command, or other executable code.
    # @option [Integer] function_group_id - Specifies the function group index if the defined name refers to a function.
    #                                       The function group defines the general category for the function.
    #                                       This attribute is used when there is an add-in or other code project associated with the file.
    #                                       See Open Office XML Part 1 for more info.
    # @option [String] short_cut_key - Specifies the keyboard shortcut for the defined name.
    # @option [Boolean] publish_to_server - Specifies a boolean value that indicates whether the defined name is included in the
    #                                       version of the workbook that is published to or rendered on a Web or application server.
    # @option [Boolean] workbook_parameter - Specifies a boolean value that indicates that the name is used as a workbook parameter on a
    #                                        version of the workbook that is published to or rendered on a Web or application server.
    def initialize(formula, options={})
      @formula = formula
      parse_options options
    end

    attr_reader :local_sheet_id

    # The local sheet index (0-based)
    # @param [Integer] value the unsigned integer index of the sheet this defined_name applies to.
    def local_sheet_id=(value)
      Axlsx::validate_unsigned_int(value)
      @local_sheet_id = value
    end

    string_attr_accessor :short_cut_key, :status_bar, :help, :description, :custom_menu, :comment, :name, :formula

    boolean_attr_accessor :workbook_parameter, :publish_to_server, :xlm, :vb_proceedure, :function, :hidden

    serializable_attributes :short_cut_key, :status_bar, :help, :description, :custom_menu, :comment,
      :workbook_parameter, :publish_to_server, :xlm, :vb_proceedure, :function, :hidden, :local_sheet_id

    def to_xml_string(str='')
      raise ArgumentError, 'you must specify the name for this defined name. Please read the documentation for Axlsx::DefinedName for more details' unless name
      str << ('<definedName ' << 'name="' << name << '" ')
      serialized_attributes str
      str << ('>' << @formula << '</definedName>')
    end
  end
end
