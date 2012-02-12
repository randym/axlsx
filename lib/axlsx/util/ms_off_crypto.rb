# encoding: UTF-8
require 'digest'
require 'base64'
require 'openssl'

module Axlsx

  # The MsOffCrypto class implements ECMA-367 encryption based on the MS-OFF-CRYPTO specification
  class MsOffCrypto

    # Creates a new MsOffCrypto Object
    # @param [String] file_name the location of the file you want to encrypt
    # @param [String] pwd the password to use when encrypting the file
    def initialize(file_name, pwd)
      self.password = pwd
      self.file_name = file_name
    end

    # Generates a new CBF file based on this instance of ms-off-crypto and overwrites the unencrypted file.
    def save
      cfb = Cbf.new(self)
      cfb.save
    end

    # returns the raw password used in encryption
    # @return [String]
    attr_reader :password

    # sets the password to be used for encryption
    # @param [String] v the password, @default 'password'
    # @return [String]
    def password=(v)
      @password = v || 'password'
    end

    # retruns the file name of the archive to be encrypted
    # @return [String]
    attr_reader :file_name

    # sets the filename
    # @return [String]
    def file_name=(v)
      #TODO verfify that the file specified exists and is an unencrypted xlsx archive
      @file_name = v
    end


    # encrypts and returns the package specified by the file name 
    # @return [String]
    def encrypted_package
      @encrypted_package ||= encrypt_package(file_name)
    end

    # returns the encryption info for this instance of ms-off-crypto
    # @return [String]
    def encryption_info 
      @encryption_info ||= create_encryption_info
    end

    # returns a random salt
    # @return [String]
    def salt
      @salt ||= Digest::SHA1.digest(rand(16**16).to_s)
    end

    # returns a random verifier
    # @return [String]
    def verifier 
      @verifier ||= rand(16**16).to_s
    end

    # returns the verifier encrytped
    # @return [String]
    def encrypted_verifier
      @encrypted_verifier ||= encrypt(verifier)
    end

    # returns the verifier hash encrypted
    # @return [String]
    def encrypted_verifier_hash
      @encrypted_verifier_hash ||= encrypt(verifier_hash)
    end

    # returns a verifier hash
    # @return [String]
    def verifier_hash
      @verifier_hash ||= create_verifier_hash
    end

    # returns an encryption key
    # @return [String]
    def key 
      @key ||= create_key
    end
    
    # size of unencrypted package? concated with encrypted package
    def encrypt_package(file_name)      
      package = File.open(file_name, 'r')
      crypt_pack = encrypt(package.read)      
      [crypt_pack.size].pack('q') + crypt_pack
    end

    # Generates an encryption info structure
    # @return [String]
    def create_encryption_info
      header = [3, 0, 2, 0] # version
      # Header flags copy
      header.concat [0x24, 0, 0, 0] #flags -- VERY UNSURE ABOUT THIS STILL
      # header.concat [0, 0, 0, 0] #unused
      header.concat [0xA4, 0, 0, 0] #length
      # Header
      header.concat  [0x24, 0, 0, 0] #flags again
      # header.concat [0, 0, 0, 0] #unused again,
      header.concat [0x0E, 0x66, 0, 0] #alg id
      header.concat [0x04, 0x80, 0, 0] #alg hash id
      header.concat [key.size, 0, 0, 0] #key size
      header.concat [0x18, 0, 0, 0] #provider type
      # header.concat [0, 0, 0, 0] #reserved 1
      # header.concat [0, 0, 0, 0] #reserved 2
      header.concat [0xA0, 0xC7, 0xDC, 0x2, 0, 0, 0, 0]
      header.concat "Microsoft Enhanced RSA and AES Cryptographic Provider (Prototype)".bytes.to_a.pack('s*').bytes.to_a
      header.concat [0, 0] #null terminator

      #Salt Size
      header.concat [salt.bytes.to_a.size].pack('l').bytes.to_a
      #Salt
      header.concat salt.bytes.to_a.pack('c*').bytes.to_a
      # encryption verifier
      header.concat encrypted_verifier.bytes.to_a.pack('c*').bytes.to_a

      # verifier hash size -- MUST BE 32 bytes
      header.concat [verifier_hash.bytes.to_a.size].pack('l').bytes.to_a

      #encryption verifier hash
      header.concat encrypted_verifier_hash.bytes.to_a.pack('c*').bytes.to_a

      header.flatten!
      header.pack('c*')
    end

    # 2.3.3
    def create_verifier_hash
      vh = Digest::SHA1.digest(verifier)
      vh << Array.new(32 - vh.size, 0).join('')
    end

    # 2.3.4.7 ECMA-376 Document Encryption Key Generation (Standard Encryption)
    def create_key
      sha = Digest::SHA1.new() << (salt + @password)
      (0..49999).each { |i| sha.update(i.to_s+sha.to_s) }
      key = sha.update(sha.to_s+'0').digest
      a = key.bytes.each_with_index.map { |item, i|  0x36 ^ item }
      x1 = Digest::SHA1.digest((a.concat Array.new(64 - key.size, 0x36)).to_s)
      a = key.bytes.each_with_index.map { |item, i| 0x5C ^ item }
      x2 = Digest::SHA1.digest( (a.concat Array.new(64 - key.size, 0x5C) ).to_s)
      x3 = x1 + x2
      x3.bytes.to_a[(0..31)].pack('c*')
    end

    # ensures that the a hashed decryption of the encryption verifier matches the decrypted verifier hash.
    # @return [Boolean]
    def verify_password
      v = Digest::SHA1.digest decrypt(@encrypted_verifier)
      vh = decrypt(@encrypted_verifier_hash)
      vh[0..15] == v[0..15]
    end

    # encrypts the data proved 
    # @param [String] data
    # @return [String] the encrypted data
    def encrypt(data)
      aes = OpenSSL::Cipher.new("AES-128-ECB")     
      aes.encrypt
      aes.key = key
      aes.update(data) << aes.final
    end

    # dencrypts the data proved 
    # @param [String] data
    # @return [String] the dencrypted data
    def decrypt(data)
      aes = OpenSSL::Cipher.new("AES-128-ECB")     
      aes.decrypt
      aes.key = key
      aes.update(data) << aes.final
    end

  end
end
