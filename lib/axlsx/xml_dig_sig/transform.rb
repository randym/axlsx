module Axlsx

  # This class represents an XML Canonicalization
  # @see Canonical XML http://www.w3.org/TR/2001/REC-xml-c14n-20010314
  class C14nTransform

    ALGORITHM = "http://www.w3.org/TR/2001/REC-xml-c14n-20010315"

    # Applies the tranform to the string provided
    # @param [String] document The xml document as a string to be tranformed
    def apply(transform)
      transform.doc = Nokogiri::XML(transform.doc).canonicalize
    end

    def to_xml_string(str)
      str << '<Transform Algorithm="' << ALGORITHM << '"/>'
    end
  end

  #The relationships transform algorithm is as follows:

  #Step 1: Process versioning instructions
  # THIS STEP IS A NOOP FOR AXSLX
  #1. The package implementer shall process the versioning instructions, considering that the only known
  #   namespace is the Relationships namespace.
  #
  #2. The package implementer shall remove all ignorable content, ignoring preservation attributes.
  #3. The package implementer shall remove all versioning instructions.

  #Step 2: Sort and filter relationships

  #1. The package implementer shall remove all namespace declarations except the Relationships namespace
  #   declaration. - OK FOR AXLSX
  #2. The package implementer shall remove the Relationships namespace prefix, if it is present.
  #   OK FOR AXLSX?
  #3. The package implementer shall sort relationship elements by Id value in lexicographical order,
  #   considering Id values as case-sensitive Unicode strings.
      # AXLSX RELATIONSHIPS ARE ALWAYS SORTED
  #4. The package implementer shall remove all Relationship elements that do not have eitheran Id value
  #   that matches any SourceId valueor a Type value that matches any SourceType value, among the
  #   SourceId and SourceType values specified in the transform definition. Producers and consumers shall
  #   compare values as case-sensitive Unicode strings. [M6.27] The resulting XML document holds all
  #   Relationship elements that either have an Id value that matches a SourceId value or a Type value that
  #   matches a SourceType value specified in the transform definition.
  #   erm..... yeah right...
  #
  # What they really mean is 
  #  <mdssi:RelationshipReference SourceId="rId3"/> these sub element references need to be
  #  included for every relationship that is to be transformed

  #Step 3: Prepare for canonicalization
  #1. The package implementer shall remove all characters between the Relationships start tag and the first
  #   Relationship start tag.
  #   AXLSX is fine here unless that includes carrage returns
  #
  #2. The package implementer shall remove any contents of the Relationship element.
  #   AXLSX is fine here
  #3. The package implementer shall remove all characters between the last Relationship end tag and the
  #   Relationships end tag.
  #   AXLSX is fine here unless that includes carrage returns
  #4. If there are no Relationship elements, the package implementer shall remove all characters between
  #   the Relationships start tag and the Relationships end tag.
  #   AXSLX should error if you are trying to sign a workbook that has no worksheets.
  
  # working example:
  #  http://eid-applet.googlecode.com/svn-history/r463/trunk/eid-applet-service-signer/src/main/java/be/fedict/eid/applet/service/signer/ooxml/RelationshipTransformService.java
  class RelationshipTransform
    ALGORITHM = "http://schemas.openxmlformats.org/package/2006/RelationshipTransform"

    def initialize
      @relationship_references = []
    end

    # unless i am misunderstanding that outstanding specification - we dont actually
    # need to do anything here except populate the TargetMode attribute when it is nill
    # just pass in obj.to_xml_string as the document
    # @param [String] document the document to transform
    # @return [String]
    def apply(reference)
      doc = Nokogiri::XML(reference.part)
      doc.xpath('xmlns:Relationships/xmlns:Relation').each do |relation|
        if relation['TargetMode'].nil?
          relation['TargetMode'] = 'Internal'
        end
      end
      reference.part = doc.to_xml
    end

    # serailize the tranformation
    # @param [String] str the string our serialization will be appended to.
    # @return [String]
    def to_xml_string(str="")
      str << "<Transform Algorithm=" << ALGORITHM << '">'
      @doc.xpath('//xmlns:Relationships/xmlns:Relationship').each do |relationship|
        str << '<mdssi:RelationshipReference SourceId="' << relationship.attr('Id') << '"/>'
      end
      str << "</Transform>"
    end
  end
end
