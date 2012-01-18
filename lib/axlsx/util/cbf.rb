module Axlsx

  # The Cfb class is a MS-OFF-CRYPTOGRAPHY specific OLE (MS-CBF) writer implementation. No attempt is made to re-invent the wheel for read/write of compound binary files.
  class Cbf

    # the serialization for the CBF FAT
    FAT_PACKING  = "s128"

    # the serialization for the MS-OFF-CRYPTO version stream
    VERSION_PACKING = 'l s30 l3'

    # The serialization for the MS-OFF-CRYPTO dataspace map stream
    DATA_SPACE_MAP_PACKING = 'l6 s16 l s25 x2'

    # The serialization for the MS-OFF-CRYPTO strong encrytion data space stream
    STRONG_ENCRYPTION_DATA_SPACE_PACKING = 'l3 s26'

    # The serialization for the MS-OFF-CRYPTO primary stream
    PRIMARY_PACKING = 'l3 s38 l s40 l3 x12 l'

    # The cutoff size that determines if a stream should be in the mini-fat or the fat
    MINI_CUTOFF = 4096

    # The serialization for CBF header
    HEADER_PACKING = "q x16 l s3 x10 l l x4 l*"

    # Creates a new Cbf object based on the ms_off_crypto object provided.
    # @param [MsOffCrypto] ms_off_crypto
    def initialize(ms_off_crypto)
      @file_name = ms_off_crypto.file_name  
      @ms_off_crypto = ms_off_crypto
      create_storages
      mini_fat_stream
      mini_fat
      fat
      header
    end

    # creates or returns the version storage
    # @return [Storage]
    def version
      @version ||= create_version
    end

    # returns the data space map storage
    # @return [Storage]
    def data_space_map
      @data_space_map ||= create_data_space_map
    end

    # returns the primary storage
    # @return [Storgae]
    def primary
      @primary ||= create_primary
    end

    # returns the summary information storage
    # @return [Storage]
    def summary_information
      @summary_information ||= create_summary_information
    end

    # returns the document summary information 
    # @return [Storage]
    def document_summary_information
      @document_summary_information ||= create_document_summary_information
    end

    # returns the stream of data allocated in the fat
    # @return [String]
    def fat_stream 
      @fat_stream ||= create_fat_stream
    end

    # returns the stream allocated in the mini fat.
    # return [String]
    def mini_fat_stream
      @mini_fat_stream ||= create_mini_fat_stream
    end
    
    # returns the mini fat
    # return [String]
    def mini_fat
      @mini_fat ||= create_mini_fat
    end

    # returns the fat
    # @return [String]
    def fat
      @fat ||= create_fat
    end
    
    # returns the CFB header
    # @return [String]
    def header
      @header ||= create_header
    end

    # returns the encryption info from the ms_off_crypt object provided during intialization
    # @return [String] encryption info
    def encryption_info
      @ms_off_crypto.encryption_info
    end

    # returns the encrypted package from the ms_off_crypt object provided during initalization
    # @return [String] encrypted package
    def encrypted_package
      @ms_off_crypto.encrypted_package
    end

    # writes the compound binary file to disk
    def save
      ole = File.open(@file_name, 'w')
      ole << header
      ole << fat
      @storages.each { |s| ole << s.to_s }
      ole << Array.new((512-(ole.pos % 512)), 0).pack('c*')
      ole << mini_fat
      ole << mini_fat_stream
      ole << fat_stream
      ole.close
    end
    
    private

    # Generates the storages required for ms-office-cryptography cfb
    def create_storages
      @storages = []
      @encryption_info = @ms_off_crypto.encryption_info
      @encrypted_package = @ms_off_crypto.encrypted_package

      @storages << Storage.new('EncryptionInfo', :data=>encryption_info, :left=>3, :right=>11) # example shows right child. do we need the summary info????
      @storages << Storage.new('EncryptedPackage', :data=>encrypted_package, :color=>Storage::COLORS[:red])
      @storages << Storage.new([6].pack("c")+"DataSpaces", :child=>5, :modified =>129685612740945580, :created=>129685612740819979)
      @storages << version
      @storages << data_space_map
      @storages << Storage.new('DataSpaceInfo', :right=>8, :child=>7, :created=>129685612740828880,:modified=>129685612740831800)
      @storages << strong_encryption_data_space
      @storages << Storage.new('TransformInfo', :color => Storage::COLORS[:red],  :child=>9, :created=>129685612740834130, :modified=>129685612740943959)
      @storages << Storage.new('StrongEncryptionTransform', :child=>10, :created=>129685612740834169, :modified=>129685612740942280)
      @storages << primary      
      @storages << summary_information
      @storages << document_summary_information

      # we do this at the end as we need to build the minifat stream to determine the size. #HOWEVER - it looks like the size should not include the padding?
      @storages.unshift Storage.new('Root Entry', :type=>Storage::TYPES[:root], :color=>Storage::COLORS[:red], :child=>1, :data => mini_fat_stream)

    end

    # generates the mini fat stream
    # @return [String]
    def create_mini_fat_stream
      mfs = []
      @storages.select{ |s| s.type == Storage::TYPES[:stream] && s.size < MINI_CUTOFF}.each_with_index do |stream, index|
        puts "#{stream.name.pack('c*')}: #{stream.data.size}"        
        mfs.concat stream.data
        mfs.concat Array.new(64 - (mfs.size % 64), 0) if mfs.size % 64 > 0        
        puts "mini fat stream size: #{mfs.size}"
      end
      mfs.concat(Array.new(512 - (mfs.size % 512), 0))
      mfs.pack 'c*'
    end

    # generates the fat stream.
    # @return [String]
    def create_fat_stream
      mfs = []
      @storages.select{ |s| s.type == Storage::TYPES[:stream] && s.size >= MINI_CUTOFF}.each_with_index do |stream, index|
        mfs.concat stream.data
        mfs.concat Array.new(512 - (mfs.size % 512), 0) if mfs.size % 512 > 0     
      end
      mfs.pack 'c*'
    end

    # creates the mini fat
    # @return [String]
    def create_mini_fat
      v_mf = []
      @storages.select{ |s| s.type == Storage::TYPES[:stream] && s.size < MINI_CUTOFF}.each do |stream|
        allocate_stream(v_mf, stream, 64)
      end
      v_mf.concat Array.new(128 - v_mf.size, -1)
      v_mf.pack 'l*'
    end

    # creates the fat
    # @return [String]
    def create_fat
      v_fat = [-3]
      # storages four per sector, allocation forces directories to start at sector ID 0
      allocate_stream(v_fat, @storages, 4)
      # fat entry for minifat
      allocate_stream(v_fat, 0, 512)
      # fat entry for minifat stream
      @storages[0].sector = v_fat.size
      allocate_stream(v_fat, mini_fat_stream, 512)
      # fat entries for encrypted package storage
      # what to do about DIFAT for larger packages...
      if @encrypted_package.size > (109 - v_fat.size) * 512
        raise ArgumentError, "Your package is too big!"
      end

      if @encrypted_package.size >= MINI_CUTOFF
        allocate_stream(v_fat, @encrypted_package, 512)
      end

      v_fat.concat Array.new(128 - v_fat.size, -1) if v_fat.size < 128 #pack in unused sectors
      v_fat.pack 'l*'
    end
    
    # Creates the version storage
    # @return [Storage]
    def create_version
      v_stream= [60, "Microsoft.Container.DataSpaces".bytes.to_a, 1, 1, 1].flatten!.pack VERSION_PACKING
      Storage.new('Version', :data=>v_stream, :size=>v_stream.size)
    end

    # returns the strong encryption data space storage
    # @return [Storgae]
    def strong_encryption_data_space
      @strong_encryption_data_space ||= create_strong_encryption_data_space
    end

    # Creates the data space map storage
    # @return [Storgae]
    def create_data_space_map
      v_stream = [8,1,104, 1,0, 32, "EncryptedPackage".bytes.to_a, 50, "StrongEncryptionDataSpace".bytes.to_a].flatten!.pack DATA_SPACE_MAP_PACKING      
      Storage.new('DataSpaceMap', :data=>v_stream, :left => 4, :right => 6, :size=>v_stream.size)
    end


    # creates the stron encryption data space storage
    # @return [Storgae]
    def create_strong_encryption_data_space
      v_stream = [8,1,50,"StrongEncryptionTransform".bytes.to_a,0].flatten.pack STRONG_ENCRYPTION_DATA_SPACE_PACKING
      Storage.new("StrongEncryptionDataSpace", :data=>v_stream, :size => v_stream.size)
    end
    
    # creates the primary storage
    # @return [Storgae]
    def create_primary
      v_stream = [88,1,76,"{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}".bytes.to_a].flatten
      v_stream.concat [78, "Microsoft.Container.EncryptionTransform".bytes.to_a,0,1,1,1,4].flatten
      v_stream = v_stream.pack PRIMARY_PACKING
      Storage.new([6].pack("c")+"Primary", :data=>v_stream)
    end


    # creates the summary information storage
    # @return [Storage]
    def create_summary_information
      v_stream = []
      v_stream.concat [0xFEFF, 0x0000, 0x030A, 0x0100, 0x0000, 0x0000, 0x0000, 0x0000]
      v_stream.concat [0x0000, 0x0000, 0x0000, 0x0000, 0x0100, 0x0000, 0xE085, 0x9FF2]
      v_stream.concat [0xF94F, 0x6810, 0xAB91, 0x0800, 0x2B27, 0xB3D9, 0x3000, 0x0000]
      v_stream.concat [0xAC00, 0x0000, 0x0700, 0x0000, 0x0100, 0x0000, 0x4000, 0x0000]
      v_stream.concat [0x0400, 0x0000, 0x4800, 0x0000, 0x0800, 0x0000, 0x5800, 0x0000]
      v_stream.concat [0x1200, 0x0000, 0x6800, 0x0000, 0x0C00, 0x0000, 0x8C00, 0x0000]
      v_stream.concat [0x0D00, 0x0000, 0x9800, 0x0000, 0x1300, 0x0000, 0xA400, 0x0000]
      v_stream.concat [0x0200, 0x0000, 0xE9FD, 0x0000, 0x1E00, 0x0000, 0x0800, 0x0000]
      v_stream.concat [0x7261, 0x6E64, 0x796D, 0x0000, 0x1E00, 0x0000, 0x0800, 0x0000]
      v_stream.concat [0x7261, 0x6E64, 0x796D, 0x0000, 0x1E00, 0x0000, 0x1C00, 0x0000]
      v_stream.concat [0x4D69, 0x6372, 0x6F73, 0x6F66, 0x7420, 0x4D61, 0x6369, 0x6E74]
      v_stream.concat [0x6F73, 0x6820, 0x4578, 0x6365, 0x6C00, 0x0000, 0x4000, 0x0000]
      v_stream.concat [0x10AC, 0x5396, 0x60BC, 0xCC01, 0x4000, 0x0000, 0x40F4, 0xFDAF]
      v_stream.concat [0x60BC, 0xCC01, 0x0300, 0x0000, 0x0100, 0x0000]

      v_stream = v_stream.pack "s*"

      Storage.new([5].pack('c')+"SummaryInformation", :data=>v_stream, :left => 2)
    end


    # creates the document summary information storage
    # @return [Storage]
    def create_document_summary_information
      v_stream = []
      v_stream.concat [0xFEFF, 0x0000, 0x030A, 0x0100, 0x0000, 0x0000, 0x0000, 0x0000]
      v_stream.concat [0x0000, 0x0000, 0x0000, 0x0000, 0x0100, 0x0000, 0x02D5, 0xCDD5]
      v_stream.concat [0x9C2E, 0x1B10, 0x9397, 0x0800, 0x2B2C, 0xF9AE, 0x3000, 0x0000]
      v_stream.concat [0xCC00, 0x0000, 0x0900, 0x0000, 0x0100, 0x0000, 0x5000, 0x0000]
      v_stream.concat [0x0F00, 0x0000, 0x5800, 0x0000, 0x1700, 0x0000, 0x6400, 0x0000]
      v_stream.concat [0x0B00, 0x0000, 0x6C00, 0x0000, 0x1000, 0x0000, 0x7400, 0x0000]
      v_stream.concat [0x1300, 0x0000, 0x7C00, 0x0000, 0x1600, 0x0000, 0x8400, 0x0000]
      v_stream.concat [0x0D00, 0x0000, 0x8C00, 0x0000, 0x0C00, 0x0000, 0x9F00, 0x0000]
      v_stream.concat [0x0200, 0x0000, 0xE9FD, 0x0000, 0x1E00, 0x0000, 0x0400, 0x0000]
      v_stream.concat [0x0000, 0x0000, 0x0300, 0x0000, 0x0000, 0x0C00, 0x0B00, 0x0000]
      v_stream.concat [0x0000, 0x0000, 0x0B00, 0x0000, 0x0000, 0x0000, 0x0B00, 0x0000]
      v_stream.concat [0x0000, 0x0000, 0x0B00, 0x0000, 0x0000, 0x0000, 0x1E10, 0x0000]
      v_stream.concat [0x0100, 0x0000, 0x0700, 0x0000, 0x5368, 0x6565, 0x7431, 0x000C]
      v_stream.concat [0x1000, 0x0002, 0x0000, 0x001E, 0x0000, 0x0013, 0x0000, 0x00E3]
      v_stream.concat [0x83AF, 0xE383, 0xBCE3, 0x82AF, 0xE382, 0xB7E3, 0x83BC, 0xE383]
      v_stream.concat [0x8800, 0x0300, 0x0000, 0x0100, 0x0000, 0x0000]
      v_stream = v_stream.pack 'c*'
      Storage.new([5].pack('c')+"DocumentSummaryInformation", :data=>v_stream)
    end

    # Creates the header 
    # @return [String]
    def create_header
      header = []
      header << -2226271756974174256 # identifier pack as q
      header << 196670 # version pack as L
      header << 65534 # byte order pack as s
      header << 9 # sector shift
      header << 6 # mini-sector shift
      header << (fat.size/512.0).ceil # this is the number of FAT sectors in the file at index 6 pack as L
      header << header.last # this is the first directory sector, index of 7 pack as L
      header << MINI_CUTOFF # minfat cutoff pack as L
      # MiniFat starts after directories
      header << (fat.size/512.0).ceil + (@storages.size/4.0).ceil # this is the sector id for the first minifat index 10 pack as L
      header << (mini_fat.size/512.0).ceil # minifat sector count index 11 pack as L
      header << -2 # the first DIFAT - set to end of chain until we exceed a single FAT pack as L
      header << 0 # number of DIFAT sectors, unless we go beyond 109 FAT sectors this will always be 0 pack as L
      header << 0 # first FAT sector defined in the DIFAT pack as L
      header.concat Array.new(108, -1) # Difat sectors pack as L108
      header.pack(HEADER_PACKING)
    end

    # Allocates sector chains in a allocation table based on the sector size and stream provided
    # If a storage obeject is provided, the starting sector value for the storage is updated based on the allocation performed here.
    # @param [Array] table Allocation table array
    # @param [Storage | String] stream
    # @param [Integer] size The cutoff size for the stream. 
    def allocate_stream(table, stream, size)
      stream.sector = table.size if stream.respond_to?(:sector)
      ((stream.size / size.to_f).ceil).times { table << table.size }
      table[table.size-1] = -2 # this is the CBF chain terminator
    end

  end
end
