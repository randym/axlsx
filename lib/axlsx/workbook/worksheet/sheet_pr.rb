module Axlsx
#<xsd:complexType name="CT_SheetPr">
#<xsd:sequence>
#￼￼￼￼<xsd:element name="tabColor" type="CT_Color" minOccurs="0" maxOccurs="1"/>
#<xsd:element name="outlinePr" type="CT_OutlinePr" minOccurs="0" maxOccurs="1"/>
#<xsd:element name="pageSetUpPr" type="CT_PageSetUpPr" minOccurs="0" maxOccurs="1"/>
#</xsd:sequence>
#￼￼￼￼￼￼￼￼￼￼￼￼￼<xsd:attribute name="syncHorizontal" type="xsd:boolean" use="optional" default="false"/>
#<xsd:attribute name="syncVertical" type="xsd:boolean" use="optional" default="false"/>
#<xsd:attribute name="syncRef" type="ST_Ref" use="optional"/>
#<xsd:attribute name="transitionEvaluation" type="xsd:boolean" use="optional" default="false"/>
#<xsd:attribute name="transitionEntry" type="xsd:boolean" use="optional" default="false"/>
#<xsd:attribute name="published" type="xsd:boolean" use="optional" default="true"/>
#<xsd:attribute name="codeName" type="xsd:string" use="optional"/>
#<xsd:attribute name="filterMode" type="xsd:boolean" use="optional" default="false"/>
#<xsd:attribute name="enableFormatConditionsCalculation" type="xsd:boolean" use="optional" default="true"/>
#</xsd:complexType>
  # The SheetPr class manages serialization fo a worksheet's sheetPr element.
  # Only fit_to_page is implemented
  class SheetPr


    # These attributes are all boolean so I'm doing a bit of a hand
    # waving magic show to set up the attriubte accessors
    #    
    BOOLEAN_ATTRIBUTES = [:sync_horizontal, 
                          :sync_vertical, 
                          :transtion_evaluation, 
                          :transition_entry,
                          :published,
                          :filter_mode,
                          :enable_format_conditions_calculation]


    # Creates a new SheetPr object
    # @param [Worksheet] worksheet The worksheet that owns this SheetPr object
    def initialize(worksheet)
      raise ArgumentError, "you must provide a worksheet" unless worksheet.is_a?(Worksheet)
      @worksheet = worksheet
    end

    # Dynamically create accessors for boolean attriubtes
    BOOLEAN_ATTRIBUTES.each do |attr|    
      class_eval %{
      # The #{attr} attribute reader
      # @return [Boolean]
      attr_reader :#{attr} 

      # The #{attr} writer
      # @param [Boolean] value The value to assign to #{attr}
      # @return [Boolean]
      def #{attr}=(value)
        Axlsx::validate_boolean(value)
        @#{attr} = value
      end
      }
    end

    # Anchor point for worksheet's window.
    # @return [String]
    attr_reader :code_name

    # Specifies a stable name of the sheet, which should not change over time,
    #  and does not change from user input. This name should be used by code 
    #  to reference a particular sheet.
    # @return [String]
    attr_reader :sync_ref

    # The worksheet these properties apply to!
    # @return [Worksheet]
    attr_reader :worksheet

    # @see code_name
    # @param [String] name
    def code_name=(name)
      @code_name = name
    end

    # @see sync_ref
    # @param [String] ref A cell reference (e.g. "A1")
    def sync_ref=(ref)
      @sync_ref = ref
    end

    # Serialize the object
    # @param [String] str serialized output will be appended to this object if provided.
    # @return [String]
    def to_xml_string(str = '')
      update_properties
      str << "<sheetPr #{serialized_attributes}>"
      page_setup_pr.to_xml_string(str)
      str << "</sheetPr>"
    end

    def page_setup_pr
      @page_setup_pr ||= PageSetUpPr.new
    end

    private

    def serialized_attributes(str = '')
      instance_values.each do |key, value| 
        unless %(worksheet page_setup_pr).include? key
          str << "#{Axlsx.camel(key, false)}='#{value}' " 
        end
      end
      str
    end

    
    def update_properties
      page_setup_pr.fit_to_page = worksheet.fit_to_page?
      self.filter_mode = worksheet.auto_filter.columns.size > 0
    end
  end


  class PageSetUpPr

    # creates a new page setup properties object
    # @param [Hash] options
    # @option [Boolean] fit_to_page Flag indicating whether the sheet displays Automatic Page Breaks.
    # @option [Boolean] auto_page_breaks Flag indicating whether the Fit to Page print option is enabled.
    def initialize(options = {})
       options.each do |key, value|
         self.send("#{key}=",value) if self.respond_to?("#{key}=")
       end
    end

    # Flag indicating whether the sheet displays Automatic Page Breaks.
    # @param [Boolean] value
    # @return [Boolean]
    def fit_to_page=(value)
      Axlsx.validate_boolean value
      @fit_to_page = value
    end

    # Flag indicating whether the Fit to Page print option is enabled.
    # @param [Boolean] value
    # @return [Boolean]
    def auto_page_breaks=(value)
      Alxsx.validate_boolean value
      @auto_page_breaks = value
    end

    # serialize to xml
    def to_xml_string(str='')
      str << '<pageSetUpPr '
      instance_values.each { |key, value| str << "#{Axlsx.camel(key, false)}='#{value}' " }
      str << '></pageSetUpPr>'
    end
  end
end
