module Axlsx

  # The manifest class manages the references for signed parts of a package
  class Manifest
    def initialize
      @references = []
    end

    attr_reader :references
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
  end
end
