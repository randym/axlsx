# encoding: UTF-8
module Axlsx

  # The Storage class represents a storage object or stream in a compound file.
  class Storage

    # Packing for the Storage when pushing an array of items into a byte stream
    # Name, name length, type, color, left sibling, right sibling, child, classid, state, created, modified, sector, size
    PACKING = "s32 s1 c2 l3 x16 x4 q2 l q"

    # storage types
    TYPES = {
      :root=>5,
      :stream=>2,
      :storage=>1
    }

    # Creates a byte string for this storage
    # @return [String] 
    def to_s
     data = [@name.concat(Array.new(32-@name.size, 0)), 
             @name_size, 
             @type, 
             @color, 
             @left, 
             @right, 
             @child, 
             @created,
             @modified, 
             @sector, 
             @size].flatten
      data.pack(PACKING)
    end

    # storage colors
    COLORS = {
      :red=>0, 
      :black=>1
    }

    # The color of this node in the directory tree. Defaults to black if not specified
    # @return [Integer] color
    attr_reader :color
    
    # Sets the color for this storage
    # @param [Integer] v Must be one of the COLORS constant hash values
    def color=(v)
      RestrictionValidator.validate "Storage.color", COLORS.values, v      
      @color = v
    end

    # The size of the name for this node.
    # interesting to see that office actually uses 'R' for the root directory and lists the size as 2 bytes - thus is it *NOT* null 
    # terminated. I am making this r/w so that I can override the size
    # @return [Integer] color
    attr_reader :name_size

    # the name of the stream    
    attr_reader :name

    # sets the name of the stream.
    # This will automatically set the name_size attribute
    # @return [String] name
    def name=(v)
      @name = v.bytes.to_a << 0
      @name_size = @name.size * 2
      @name
    end

    # The size of the stream
    attr_reader :size

    # The stream associated with this storage
    attr_reader :data

    # Set the data associated with the stream. If the stream type is undefined, we automatically specify the storage as a stream type.    # with the exception of storages that are type root, all storages with data should be type stream.
    # @param [String] v The data for this storages stream
    # @return [Array]
    def data=(v)
      Axlsx::validate_string(v)
      self.type = TYPES[:stream] unless @type
      @size = v.size
      @data = v.bytes.to_a
    end

    # The starting sector for the stream. If this storage is not a stream, or the root node this is nil
    # @return [Integer] sector
    attr_accessor :sector

    # The 0 based index in the directoies chain for this the left sibling of this storage. 
    
    # @return [Integer] left
    attr_accessor :left 

    # The 0 based index in the directoies chain for this the right sibling of this storage. 
    # @return [Integer] right
    attr_accessor :right 

    # The 0 based index in the directoies chain for the child of this storage. 
    # @return [Integer] child
    attr_accessor :child

    # The created attribute for the storage
    # @return [Integer] created
    attr_accessor :created

    # The modified attribute for the storage
    # @return [Integer] modified
    attr_accessor :modified

    # The type of storage
    # see TYPES
    # @return [Integer] type
    attr_reader :type

    # Sets the type for this storage. 
    # @param [Integer] v the type to specify must be one of the TYPES constant hash values. 
    def type=(v)
      RestrictionValidator.validate "Storage.type", TYPES.values, v      
      @type = v
    end

    # Creates a new storage object. 
    # @param [String] name the name of the storage
    # @option options [Integer] color @default black
    # @option options [Integer] type @default storage
    # @option options [String] data 
    # @option options [Integer] left @default -1
    # @option options [Integer] right @default -1
    # @option options [Integer] child @default -1
    # @option options [Integer] created @default 0
    # @option options [Integer] modified @default 0
    # @option options [Integer] sector @default 0
    def initialize(name, options= {})
      @left = @right = @child = -1
      @sector = @size = @created = @modified = 0
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
      @color ||= COLORS[:black]
      @type ||= (data.nil? ? TYPES[:storage] : TYPES[:stream])
      self.name = name
    end
 
  end
end
