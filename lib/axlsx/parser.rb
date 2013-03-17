module Axlsx
  class Parser
    require 'zip/zipfilesystem' 
    class << self
      def open(filename)
        raise ArgumentError unless filename
        raise Errno::ENOENT unless File.exists?(filename)
        raise Errno::EISDIR if File.directory?(filename)
        parser = Parser.new unzip(filename)
        yield parser if block_given?
        parser
      end
      def unzip(filename)
        Zip::ZipFile.open(filename)
      end
    end

    def initialize(archive)
      @archive = archive
      extract_parts
    end

    attr_reader :archive

    def extract_parts
      package.core.parse_xml(parts[:core])
      package.app.parse_xml(parts[:app])
    end

    def package
      @package ||= Axlsx::Package.new
    end

    def extract_part(part)
      Nokogiri::XML(archive.find_entry(part).get_input_stream.read)
    end

    def parts
      @parts ||=  {rels: nil,
                   styles: nil,
                   core: extract_part(CORE_PN),
                   app: extract_part(APP_PN),
                   workbook: nil,
                   content_types: nil,
                   workbook: nil,
                   shared_strings: nil,
                   drawings: [],
                   tables: [],
                   pivot_tables: [],
                   comments: [],
                   charts: [],
                   images: [],
                   worksheets: []}
    end
  end
end
