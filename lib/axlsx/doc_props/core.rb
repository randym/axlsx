# encoding: UTF-8
module Axlsx

  # The core object for the package.
  # @note Packages manage their own core object.
  # @see Package#core
  class Core
 
    # Creates a new Core object.
    # @option options [String] creator
    # @option options [Time] created
    def initialize(options={})
      @creator = options[:creator] || 'axlsx'
      @created = options[:created]
    end

    # The author of the document. By default this is 'axlsx'
    # @return [String]
    attr_accessor :creator

    # Creation time of the document. If nil, the current time will be used.
    attr_accessor :created

    # serializes the core.xml document
    # @return [String]
    def to_xml_string(str = '')
      str << '<?xml version="1.0" encoding="UTF-8"?>'
      str << ('<cp:coreProperties xmlns:cp="' << CORE_NS << '" xmlns:dc="' << CORE_NS_DC << '" ')
      str << ('xmlns:dcmitype="' << CORE_NS_DCMIT << '" xmlns:dcterms="' << CORE_NS_DCT << '" ')
      str << ('xmlns:xsi="' << CORE_NS_XSI << '">')
      str << ('<dc:creator>' << self.creator << '</dc:creator>')
      str << ('<dcterms:created xsi:type="dcterms:W3CDTF">' << (created || Time.now).strftime('%Y-%m-%dT%H:%M:%S') << 'Z</dcterms:created>')
      str << '<cp:revision>0</cp:revision>'
      str << '</cp:coreProperties>'
    end

  end

end
