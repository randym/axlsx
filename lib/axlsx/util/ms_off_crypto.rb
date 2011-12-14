require 'digest'
require 'base64'
require 'openssl'
module Axlsx
  class MsOffCrypto

    attr_reader :verifier
    attr_reader :key

    def initialize(password = "passowrd")
      @password = password
      @salt_size = 0x10
      @key_size = 0x100
      @verifier = rand(16**16).to_s 

      #fixed salt for testing
      @salt = [0x90,0xAC,0x68,0x0E,0x76,0xF9,0x43,0x2B,0x8D,0x13,0xB7,0x1D,0xB7,0xC0,0xFC,0x0D].join
      # @salt =Digest::SHA1.digest(rand(16**16).to_s)
    end

    def encryption_info
      #         v.major v.minor flags           header length   flags           size            # AES 128 bit
      header = [3, 0, 2, 0, 0x24, 0, 0, 0, 0xA4, 0, 0, 0, 0x24, 0, 0, 0, 0, 0, 0, 0, 0x0E, 0x66, 0, 0]
      header.concat [0x04, 0x80, 0, 0, 0x80, 0, 0, 0, 0x18, 0, 0, 0, 0xA0, 0xC7, 0xDC, 0x2, 0, 0, 0, 0]
      header.concat "Microsoft Enhanced RSA and AES Cryptographic Provider (Prototype)".bytes.to_a.pack('s*').bytes.to_a
      header.concat [0, 0]
      header.concat [0x10, 0, 0, 0]
      header.concat [0x90,0xAC,0x68,0x0E,0x76,0xF9,0x43,0x2B,0x8D,0x13,0xB7,0x1D,0xB7,0xC0,0xFC,0x0D]
      header.concat encrypted_verifier.bytes.to_a.pack('c*').bytes.to_a
      header.concat [20, 0,0,0]
      header.concat encrypted_verifier_hash.bytes.to_a.pack('c*').bytes.to_a
      header.flatten!
      header.pack('c*')
    end

    def encryption_verifier
      {:salt_size => @salt_size, 
       :salt => @salt, 
       :encrypted_verifier => encrypted_verifier,
       :varifier_hash_size => 0x14,
        :encrypted_verifier_hash =>  encrypted_verifier_hash}
    end

    # 2.3.3
    def encrypted_verifier
      @encrypted_verifier ||= encrypt(@verifier)
    end

    # 2.3.3
    def encrypted_verifier_hash
      verifier_hash = Digest::SHA1.digest(@verifier)
      verifier_hash << Array.new(32 - verifier_hash.size, 0).join('')
      @encrypted_verifier_hash ||= encrypt(verifier_hash)
    end

    # 2.3.4.7 ECMA-376 Document Encryption Key Generation (Standard Encryption)
    def key
      sha = Digest::SHA1.new() << (@salt + @password)
      (0..49999).each { |i| sha.update(i.to_s+sha.to_s) }
      key = sha.update(sha.to_s+'0').digest
      a = key.bytes.each_with_index.map { |item, i|  0x36 ^ item }
      x1 = Digest::SHA1.digest((a.concat Array.new(64 - key.size, 0x36)).to_s)
      a = key.bytes.each_with_index.map { |item, i| 0x5C ^ item }
      x2 = Digest::SHA1.digest( (a.concat Array.new(64 - key.size, 0x5C) ).to_s)
      x3 = x1 + x2
      @key ||= x3.bytes.to_a[(0..31)].pack('c*')
    end

    def verify_password
      puts decrypt(@encrypted_verifier)
    end

    def encrypt(data)
      aes = OpenSSL::Cipher.new("AES-128-ECB")     
      aes.encrypt
      aes.key = key
      aes.update(data)
    end

    def decrypt(data)
      aes = OpenSSL::Cipher.new("AES-128-ECB")     
      aes.decrypt
      aes.key = key
      aes.update(data)
    end

  end
end
