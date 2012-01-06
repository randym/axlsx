module Axlsx

  # The Cfb class is a MS-OFF-CRYPTOGRAPHY specific OLE (MS-CBF) writer implementation. No attempt is made to re-invent the wheel for read/write of compound binary files.
  class Cbf

    # the serialization for the CBF FAT
    FAT_PACKING  = "s128"

    # the serialization for the MS-OFF-CRYPTO version stream
    VERSION_PACKING = 'l s30 l3'

    # The serialization for the MS-OFF-CRYPTO dataspace map stream
    DATA_SPACE_MAP_PACKING = 'l6 s16 l s25'

    # The serialization for the MS-OFF-CRYPTO strong encrytion data space stream
    STRONG_ENCRYPTION_DATA_SPACE_PACKING = 'l3 s25'

    # The serialization for the MS-OFF-CRYPTO primary stream
    PRIMARY_PACKING = 'l3 s38 l s39 l3 x12 l x2'

    # The cutoff size that determines if a stream should be in the mini-fat or the fat
    MINI_CUTOFF = 4096

    # The serialization for CBF header
    HEADER_PACKING = "q x16 l s3 x10 l l x4 l*"

    # Creates a new Cbf object based on the ms_off_crypto object provided.
    # @param [MsOffCrypto] ms_off_crypto
    def initialize(ms_off_crypto)
      @file_name = ms_off_crypto.file_name  
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
      @encryption_info = ms_off_crypto.encryption_info
      @encrypted_package = ms_off_crypto.encrypted_package
      @storages << Storage.new('R', :type=>Storage::TYPES[:root], :color=>Storage::COLORS[:red], :child=>1, :modified=>129685612742510730)
      @storages.last.name_size = 2
      @storages << Storage.new('EncryptionInfo', :data=>@encryption_info, :left=>3, :size => @encryption_info.size) # example shows right child. do we need the summary info????
      @storages << Storage.new('EncryptedPackage', :data=>@encrypted_package, :color=>Storage::COLORS[:red], :size=>@encrypted_package.size)
      @storages << Storage.new([6].pack("c")+"DataSpaces", :child=>5, :modified =>129685612740945580, :created=>129685612740819979)
      @storages << version
      @storages << data_space_map
      @storages << Storage.new('DataSpaceInfo', :right=>8, :child=>7, :created=>129685612740828880,:modified=>129685612740831800)
      @storages << strong_encryption_data_space
      @storages << Storage.new('TransformInfo', :color => Storage::COLORS[:red],  :child=>9, :created=>129685612740834130, :modified=>129685612740943959)
      @storages << Storage.new('StrongEncryptionTransform', :child=>10, :created=>129685612740834169, :modified=>129685612740942280)
      @storages << primary      
    end

    # generates the mini fat stream
    # @return [String]
    def create_mini_fat_stream
      mfs = []
      @storages.select{ |s| s.type == Storage::TYPES[:stream] && s.size < MINI_CUTOFF}.each_with_index do |stream, index|
        mfs.concat stream.data
        mfs.concat Array.new(64 - (mfs.size % 64), 0) if mfs.size % 64        
      end
      @storages[0].size = mfs.size
      mfs.concat(Array.new(512 - (mfs.size % 512), 0))
      mfs.pack 'c*'
    end

    # generates the fat stream.
    # @return [String]
    def create_fat_stream
      mfs = []
      @storages.select{ |s| s.type == Storage::TYPES[:stream] && s.size >= MINI_CUTOFF}.each_with_index do |stream, index|
        mfs.concat stream.data
        mfs.concat Array.new(512 - (mfs.size % 512), 0) if mfs.size % 512        
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
      v_stream = [8,1,50,"StrongEncryptionTransform".bytes.to_a].flatten.pack STRONG_ENCRYPTION_DATA_SPACE_PACKING
      Storage.new("StrongEncryptionDataSpace", :data=>v_stream, :size => v_stream.size)
    end
    
    # creates the primary storage
    # @return [Storgae]
    def create_primary
      v_stream = [88,1,76,"{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}".bytes.to_a].flatten
      v_stream.concat [78, "Microsoft.Container.EncryptionTransform".bytes.to_a,1,1,1,4].flatten
      v_stream = v_stream.pack PRIMARY_PACKING
      Storage.new([6].pack("c")+"Primary", :data=>v_stream)
    end


    SUMMARY_INFORMATION_PACKING = ""
    # creates the summary information storage
    # @return [Storage]
    def create_summary_information
      # FEFF 0000 030A 0100 0000 0000 0000 0000
      # 0000 0000 0000 0000 0100 0000 E085 9FF2 
      # F94F 6810 AB91 0800 2B27 B3D9 3000 0000
      # AC00 0000 0700 0000 0100 0000 4000 0000
      # 0400 0000 4800 0000 0800 0000 5800 0000
      # 1200 0000 6800 0000 0C00 0000 8C00 0000
      # 0D00 0000 9800 0000 1300 0000 A400 0000 
      # 0200 0000 E9FD 0000 1E00 0000 0800 0000 
      # 7261 6E64 796D 0000 1E00 0000 0800 0000 
      # 7261 6E64 796D 0000 1E00 0000 1C00 0000 
      # 4D69 6372 6F73 6F66 7420 4D61 6369 6E74 
      # 6F73 6820 4578 6365 6C00 0000 4000 0000 
      # 10AC 5396 60BC CC01 4000 0000 40F4 FDAF
      # 60BC CC01 0300 0000 0100 0000
      v_stream = []
      v_stream = v_stream.pack SUMMARY_INFORMATION_PACKING
      Storage.new([5].pack('c')+"SummaryInformation", :data=>v_stream)
    end

    DOCUMENT_SUMMARY_INFORMATION_PACKING = ""
    # creates the document summary information storage
    # @return [Storage]
    def create_document_summary_information
      v_stream = []
      v_stream = v_stream.pack DOCUMENT_SUMMARY_INFORMATION_PACKING
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
