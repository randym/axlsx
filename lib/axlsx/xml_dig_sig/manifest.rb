module Axlsx

  # The manifest class manages the references for signed parts of a package
  class Manifest
    def initialize
      @relations = []
    end

    attr_reader :relations
    def add_relation(package_part)
      # based on the package part content type create a new relation with the appropriate transforms and add it
      # to the manifiest reference collection
      
    end
  end
end
