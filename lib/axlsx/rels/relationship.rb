# encoding: UTF-8
module Axlsx
  # A relationship defines a reference between package parts.
  # @note Packages automatically manage relationships.
  class Relationship
    
    class << self
      # Keeps track of all instances of this class.
      # @return [Array]
      def instances
        @instances ||= []
      end
      
      # Clear cached instances.
      # 
      # This should be called before serializing a package (see {Package#serialize} and
      # {Package#to_stream}) to make sure that serialization is idempotent (i.e. 
      # Relationship instances are generated with the same IDs everytime the package
      # is serialized).
      # 
      # Also, calling this avoids memory leaks (cached instances lingering around 
      # forever). 
      def clear_cached_instances
        @instances = []
      end
      
      # Generate and return a unique id (eg. `rId123`) Used for setting {#Id}. 
      #
      # The generated id depends on the number of cached instances, so using
      # {clear_cached_instances} will automatically reset the generated ids, too.
      # @return [String]
      def next_free_id
        "rId#{@instances.size + 1}"
      end
    end

    # The id of the relationship (eg. "rId123"). Most instances get their own unique id. 
    # However, some instances need to share the same id – see {#should_use_same_id_as?}
    # for details.
    # @return [String]
    attr_reader :Id
    
    # The location of the relationship target
    # @return [String]
    attr_reader :Target

    # The type of relationship
    # @note Supported types are defined as constants in Axlsx:
    # @see XML_NS_R
    # @see TABLE_R
    # @see PIVOT_TABLE_R
    # @see WORKBOOK_R
    # @see WORKSHEET_R
    # @see APP_R
    # @see RELS_R
    # @see CORE_R
    # @see STYLES_R
    # @see CHART_R
    # @see DRAWING_R
    # @return [String]
    attr_reader :Type

    # The target mode of the relationship
    # used for hyperlink type relationships to mark the relationship to an external resource
    # TargetMode can be specified during initialization by passing in a :target_mode option
    # Target mode must be :external for now.
    attr_reader :TargetMode

    # The source object the relations belongs to (e.g. a hyperlink, drawing, ...). Needed when
    # looking up the relationship for a specific object (see {Relationships#for}).
    attr_reader :source_obj
    
    # Initializes a new relationship. 
    # @param [Object] source_obj see {#source_obj}
    # @param [String] type The type of the relationship
    # @param [String] target The target for the relationship
    # @option [Symbol] :target_mode only accepts :external.
    def initialize(source_obj, type, target, options={})
      @source_obj = source_obj
      self.Target=target
      self.Type=type
      self.TargetMode = options[:target_mode] if options[:target_mode]
      @Id = if (existing = self.class.instances.find{ |i| should_use_same_id_as?(i) })
        existing.Id
      else
        self.class.next_free_id
      end
      self.class.instances << self
    end

    # @see Target
    def Target=(v) Axlsx::validate_string v; @Target = v end
    # @see Type
    def Type=(v) Axlsx::validate_relationship_type v; @Type = v end

    # @see TargetMode
    def TargetMode=(v) RestrictionValidator.validate 'Relationship.TargetMode', [:External, :Internal], v; @TargetMode = v; end

    # serialize relationship
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      h = self.instance_values.reject{|k, _| k == "source_obj"}
      str << '<Relationship '
      str << h.map { |key, value| '' << key.to_s << '="' << Axlsx::coder.encode(value.to_s) << '"'}.join(' ')
      str << '/>'
    end
    
    # Whether this relationship should use the same id as `other`.
    #
    # Instances designating the same relationship need to use the same id. We can not simply
    # compare the {#Target} attribute, though: `foo/bar.xml`, `../foo/bar.xml`, 
    # `../../foo/bar.xml` etc. are all different but probably mean the same file (this 
    # is especially an issue for relationships in the context of pivot tables). So lets
    # just ignore this attribute for now (except when {#TargetMode} is set to `:External` –
    # then {#Target} will be an absolute URL and thus can safely be compared).
    #
    # @todo Implement comparison of {#Target} based on normalized path names.
    # @param other [Relationship]
    def should_use_same_id_as?(other)
      result = self.source_obj == other.source_obj && self.Type == other.Type && self.TargetMode == other.TargetMode
      if self.TargetMode == :External
        result &&= self.Target == other.Target
      end
      result
    end
    
  end
end
