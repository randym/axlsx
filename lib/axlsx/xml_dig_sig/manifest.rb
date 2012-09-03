module Axlsx

  # The manifest class manages the references for signed parts of a
  # packages. @note The recommended way to manage signatures is via Package#sign
  # You should not be using this class directly.
  class Manifest

    # @TODO add Themes if we ever implement them.
    SIGNABLE_CONTENT_TYPES = [RELS_CT, STYLES_CT, WORKBOOK_CT, DRAWING_CT, TABLE_CT, COMMENT_CT, CHART_CT, SHARED_STRINGS_CT, WORKSHEET_CT]


    # Creates a new Manifest instance
    def initialize
      @references = []
    end

    def is_signable_part(part)
      SIGNABLE_CONETNET_TYPES.inclue? part.content_type
    end

    # The references for this manifest. 
    # A reference is the base elememt of what will be signed in the workbook.
    # @return [Array]
    attr_reader :references

    # Add a reference to the references collection.
    # @params [Hahs] package_part This is a part from an Axlsx::Package.
    # It must contain a minimum of type and doc and entity members.
    # @note This method manages the appropriate transforms for the part 
    # based on the content type specified. To whit, relations have the
    # RelationshipTransform and C14n Transform applied. Other types have 
    # no transforms.
    #
    # @see Axlsx::Package#parts
    #
    # @return [Array]
    def add_reference_for_part(package_part)
      # based on the package part content type create a new reference 
      # with the appropriate transforms and add it
      # to the manifiest reference collection.
      reference = Reference.new(package_part.doc, "/#{package_part.entity}?ContentType=#{package_part.content_type}")
      if package_part.content_type == RELS_CT
        reference.add_transform RelationshipTransform.new
        reference.add_transfrom C14nTransform.new 
      end
      @references << reference
    end

    # Serialize the manifiest
    # @param [String] str The string to append this manifest's content to.
    # @return [String]
    def to_xml_string(str="")
      str << '<Manifest>'
      references.each { |reference| reference.to_xml_string(str) }
      str << '</Manifest>'
    end
  end
end
