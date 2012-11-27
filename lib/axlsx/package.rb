# encoding: UTF-8
module Axlsx
  # Package is responsible for managing all the bits and peices that Open Office XML requires to make a valid
  # xlsx document including valdation and serialization.
  class Package
    include Axlsx::OptionsParser

    # provides access to the app doc properties for this package
    # see App
    attr_reader :app

    # provides access to the core doc properties for the package
    # see Core
    attr_reader :core

    # Initializes your package
    #
    # @param [Hash] options A hash that you can use to specify the author and workbook for this package.
    # @option options [String] :author The author of the document
    # @option options [Boolean] :use_shared_strings This is passed to the workbook to specify that shared strings should be used when serializing the package.
    # @example Package.new :author => 'you!', :workbook => Workbook.new
    def initialize(options={})
      @workbook = nil
      @core, @app  =  Core.new, App.new
      @core.creator = options[:author] || @core.creator
      parse_options options
      yield self if block_given?
    end

    # Shortcut to specify that the workbook should use autowidth
    # @see Workbook#use_autowidth
    def use_autowidth=(v)
      Axlsx::validate_boolean(v);
      workbook.use_autowidth = v
    end


    # Shortcut to specify that the workbook should use shared strings
    # @see Workbook#use_shared_strings
    def use_shared_strings=(v)
      Axlsx::validate_boolean(v);
      workbook.use_shared_strings = v
    end

    # Shortcut to determine if the workbook is configured to use shared strings
    # @see Workbook#use_shared_strings
    def use_shared_strings
      workbook.use_shared_strings
    end

    # The workbook this package will serialize or validate.
    # @return [Workbook] If no workbook instance has been assigned with this package a new Workbook instance is returned.
    # @raise ArgumentError if workbook parameter is not a Workbook instance.
    # @note As there are multiple ways to instantiate a workbook for the package,
    #   here are a few examples:
    #     # assign directly during package instanciation
    #     wb = Package.new(:workbook => Workbook.new).workbook
    #
    #     # get a fresh workbook automatically from the package
    #     wb = Pacakge.new().workbook
    #     #     # set the workbook after creating the package
    #     wb = Package.new().workbook = Workbook.new
    def workbook
      @workbook || @workbook = Workbook.new
      yield @workbook if block_given?
      @workbook
    end

    #def self.parse(input, confirm_valid = false)
    #  p = Package.new
    #  z = Zip::ZipFile.open(input)
    #  p.workbook = Workbook.parse z.get_entry(WORKBOOK_PN)
    #  p
    #end

    # @see workbook
    def workbook=(workbook) DataTypeValidator.validate "Package.workbook", Workbook, workbook; @workbook = workbook; end

    # Serialize your workbook to disk as an xlsx document.
    #
    # @param [String] output The name of the file you want to serialize your package to
    # @param [Boolean] confirm_valid Validate the package prior to serialization.
    # @return [Boolean] False if confirm_valid and validation errors exist. True if the package was serialized
    # @note A tremendous amount of effort has gone into ensuring that you cannot create invalid xlsx documents.
    #   confirm_valid should be used in the rare case that you cannot open the serialized file.
    # @see Package#validate
    # @example
    #   # This is how easy it is to create a valid xlsx file. Of course you might want to add a sheet or two, and maybe some data, styles and charts.
    #   # Take a look at the README for an example of how to do it!
    #
    #   #serialize to a file
    #   p = Axlsx::Package.new
    #   # ......add cool stuff to your workbook......
    #   p.serialize("example.xlsx")
    #
    #   # Serialize to a stream
    #   s = p.to_stream()
    #   File.open('example_streamed.xlsx', 'w') { |f| f.write(s.read) }
    def serialize(output, confirm_valid=false)
      return false unless !confirm_valid || self.validate.empty?
      Zip::ZipOutputStream.open(output) do |zip|
        write_parts(zip)
      end
      true
    end


    # Serialize your workbook to a StringIO instance
    # @param [Boolean] confirm_valid Validate the package prior to serialization.
    # @return [StringIO|Boolean] False if confirm_valid and validation errors exist. rewound string IO if not.
    def to_stream(confirm_valid=false)
      return false unless !confirm_valid || self.validate.empty?
      zip = write_parts(Zip::ZipOutputStream.new("streamed", true))
      stream = zip.close_buffer
      stream.rewind
      stream
    end

    # Encrypt the package into a CFB using the password provided
    # This is not ready yet
    def encrypt(file_name, password)
      return false
      # moc = MsOffCrypto.new(file_name, password)
      # moc.save
    end

    # Validate all parts of the package against xsd schema.
    # @return [Array] An array of all validation errors found.
    # @note This gem includes all schema from OfficeOpenXML-XMLSchema-Transitional.zip and OpenPackagingConventions-XMLSchema.zip
    #   as per ECMA-376, Third edition. opc schema require an internet connection to import remote schema from dublin core for dc,
    #   dcterms and xml namespaces. Those remote schema are included in this gem, and the original files have been altered to
    #   refer to the local versions.
    #
    #   If by chance you are able to creat a package that does not validate it indicates that the internal
    #   validation is not robust enough and needs to be improved. Please report your errors to the gem author.
    # @see http://www.ecma-international.org/publications/standards/Ecma-376.htm
    # @example
    #  # The following will output any error messages found in serialization.
    #  p = Axlsx::Package.new
    #  # ... code to create sheets, charts, styles etc.
    #  p.validate.each { |error| puts error.message }
    def validate
      errors = []
      parts.each do |part|
        errors.concat validate_single_doc(part[:schema], part[:doc]) unless part[:schema].nil?
      end
      errors
    end

    private

    # Writes the package parts to a zip archive.
    # @param [Zip::ZipOutputStream] zip
    # @return [Zip::ZipOutputStream]
    def write_parts(zip)
      p = parts
      p.each do |part|
        unless part[:doc].nil?
          zip.put_next_entry(part[:entry])
          entry = ['1.9.2', '1.9.3'].include?(RUBY_VERSION) ? part[:doc].force_encoding('BINARY') : part[:doc]
          zip.puts(entry)
        end
        unless part[:path].nil?
          zip.put_next_entry(part[:entry]);
          # binread for 1.9.3
          zip.write IO.respond_to?(:binread) ? IO.binread(part[:path]) : IO.read(part[:path])
        end
      end
      zip
    end

    # The parts of a package
    # @return [Array] An array of hashes that define the entry, document and schema for each part of the package.
    # @private
    def parts
      parts = [
       {:entry => RELS_PN, :doc => relationships.to_xml_string, :schema => RELS_XSD},
       {:entry => "xl/#{STYLES_PN}", :doc => workbook.styles.to_xml_string, :schema => SML_XSD},
       {:entry => CORE_PN, :doc => @core.to_xml_string, :schema => CORE_XSD},
       {:entry => APP_PN, :doc => @app.to_xml_string, :schema => APP_XSD},
       {:entry => WORKBOOK_RELS_PN, :doc => workbook.relationships.to_xml_string, :schema => RELS_XSD},
       {:entry => CONTENT_TYPES_PN, :doc => content_types.to_xml_string, :schema => CONTENT_TYPES_XSD},
       {:entry => WORKBOOK_PN, :doc => workbook.to_xml_string, :schema => SML_XSD}
      ]

      workbook.drawings.each do |drawing|
        parts << {:entry => "xl/#{drawing.rels_pn}", :doc => drawing.relationships.to_xml_string, :schema => RELS_XSD}
        parts << {:entry => "xl/#{drawing.pn}", :doc => drawing.to_xml_string, :schema => DRAWING_XSD}
      end


      workbook.tables.each do |table|
        parts << {:entry => "xl/#{table.pn}", :doc => table.to_xml_string, :schema => SML_XSD}
      end
      workbook.pivot_tables.each do |pivot_table|
        cache_definition = pivot_table.cache_definition
        parts << {:entry => "xl/#{pivot_table.rels_pn}", :doc => pivot_table.relationships.to_xml_string, :schema => RELS_XSD}
        parts << {:entry => "xl/#{pivot_table.pn}", :doc => pivot_table.to_xml_string} #, :schema => SML_XSD}
        parts << {:entry => "xl/#{cache_definition.pn}", :doc => cache_definition.to_xml_string} #, :schema => SML_XSD}
      end

      workbook.comments.each do|comment|
        if comment.size > 0
          parts << { :entry => "xl/#{comment.pn}", :doc => comment.to_xml_string, :schema => SML_XSD }
          parts << { :entry => "xl/#{comment.vml_drawing.pn}", :doc => comment.vml_drawing.to_xml_string, :schema => nil }
        end
      end

      workbook.charts.each do |chart|
        parts << {:entry => "xl/#{chart.pn}", :doc => chart.to_xml_string, :schema => DRAWING_XSD}
      end

      workbook.images.each do |image|
        parts << {:entry => "xl/#{image.pn}", :path => image.image_src}
      end

      if use_shared_strings
        parts << {:entry => "xl/#{SHARED_STRINGS_PN}", :doc => workbook.shared_strings.to_xml_string, :schema => SML_XSD}
      end

      workbook.worksheets.each do |sheet|
        parts << {:entry => "xl/#{sheet.rels_pn}", :doc => sheet.relationships.to_xml_string, :schema => RELS_XSD}
        parts << {:entry => "xl/#{sheet.pn}", :doc => sheet.to_xml_string, :schema => SML_XSD}
      end
      parts
    end

    # Performs xsd validation for a signle document
    #
    # @param [String] schema path to the xsd schema to be used in validation.
    # @param [String] doc The xml text to be validated
    # @return [Array] An array of all validation errors encountered.
    # @private
    def validate_single_doc(schema, doc)
      schema = Nokogiri::XML::Schema(File.open(schema))
      doc = Nokogiri::XML(doc)
      errors = []
      schema.validate(doc).each do |error|
        errors << error
      end
      errors
    end

    # Appends override objects for drawings, charts, and sheets as they exist in your workbook to the default content types.
    # @return [ContentType]
    # @private
    def content_types
      c_types = base_content_types
      workbook.drawings.each do |drawing|
        c_types << Axlsx::Override.new(:PartName => "/xl/#{drawing.pn}",
                                       :ContentType => DRAWING_CT)
      end

      workbook.charts.each do |chart|
        c_types << Axlsx::Override.new(:PartName => "/xl/#{chart.pn}",
                                       :ContentType => CHART_CT)
      end

      workbook.tables.each do |table|
        c_types << Axlsx::Override.new(:PartName => "/xl/#{table.pn}",
                                       :ContentType => TABLE_CT)
      end

      workbook.pivot_tables.each do |pivot_table|
        c_types << Axlsx::Override.new(:PartName => "/xl/#{pivot_table.pn}",
                                       :ContentType => PIVOT_TABLE_CT)
        c_types << Axlsx::Override.new(:PartName => "/xl/#{pivot_table.cache_definition.pn}",
                                       :ContentType => PIVOT_TABLE_CACHE_DEFINITION_CT)
      end

      workbook.comments.each do |comment|
        if comment.size > 0
        c_types << Axlsx::Override.new(:PartName => "/xl/#{comment.pn}",
                                       :ContentType => COMMENT_CT)
        end
      end

      if workbook.comments.size > 0
        c_types << Axlsx::Default.new(:Extension => "vml", :ContentType => VML_DRAWING_CT)
      end

      workbook.worksheets.each do |sheet|
        c_types << Axlsx::Override.new(:PartName => "/xl/#{sheet.pn}",
                                         :ContentType => WORKSHEET_CT)
      end
      exts = workbook.images.map { |image| image.extname }
      exts.uniq.each do |ext|
        ct = if  ['jpeg', 'jpg'].include?(ext)
               JPEG_CT
             elsif ext == 'gif'
               GIF_CT
             elsif ext == 'png'
               PNG_CT
             end
        c_types << Axlsx::Default.new(:ContentType => ct, :Extension => ext )
      end
      if use_shared_strings
        c_types << Axlsx::Override.new(:PartName => "/xl/#{SHARED_STRINGS_PN}",
                                       :ContentType => SHARED_STRINGS_CT)
      end
      c_types
    end

    # Creates the minimum content types for generating a valid xlsx document.
    # @return [ContentType]
    # @private
    def base_content_types
      c_types = ContentType.new()
      c_types <<  Default.new(:ContentType => RELS_CT, :Extension => RELS_EX)
      c_types <<  Default.new(:Extension => XML_EX, :ContentType => XML_CT)
      c_types << Override.new(:PartName => "/#{APP_PN}", :ContentType => APP_CT)
      c_types << Override.new(:PartName => "/#{CORE_PN}", :ContentType => CORE_CT)
      c_types << Override.new(:PartName => "/xl/#{STYLES_PN}", :ContentType => STYLES_CT)
      c_types << Axlsx::Override.new(:PartName => "/#{WORKBOOK_PN}", :ContentType => WORKBOOK_CT)
      c_types.lock
      c_types
    end

    # Creates the relationships required for a valid xlsx document
    # @return [Relationships]
    # @private
    def relationships
      rels = Axlsx::Relationships.new
      rels << Relationship.new(WORKBOOK_R, WORKBOOK_PN)
      rels << Relationship.new(CORE_R, CORE_PN)
      rels << Relationship.new(APP_R, APP_PN)
      rels.lock
      rels
    end
  end
end

