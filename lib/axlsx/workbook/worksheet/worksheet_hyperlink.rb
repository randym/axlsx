module Axlsx

  # A worksheet hyperlink object. Note that this is not the same as a drawing hyperlink object.
  class WorksheetHyperlink

    include Axlsx::OptionsParser
    include Axlsx::Accessors
    include Axlsx::SerializedAttributes
    # Creates a new hyperlink object.
    # @note the preferred way to add hyperlinks to your worksheet is the Worksheet#add_hyperlink method
    # @param [Worksheet] worksheet the Worksheet that owns this hyperlink
    # @param [Hash] options options to use when creating this hyperlink
    # @option [String] display Display string, if different from string in string table. This is a property on the hyperlink object, but does not need to appear in the spreadsheet application UI.
    # @option [String] location Location within target. If target is a workbook (or this workbook) this shall refer to a sheet and cell or a defined name. Can also be an HTML anchor if target is HTML file.
    # @option [String] tooltip The tip to display when the user positions the mouse cursor over this hyperlink
    # @option [Symbol] target This is :external by default. If you set it to anything else, the location is interpreted to be the current workbook.
    # @option [String|Cell] ref The location of this hyperlink in the worksheet
    def initialize(worksheet, options={})
      DataTypeValidator.validate "Hyperlink.worksheet", [Worksheet], worksheet
      @worksheet = worksheet
      @target = :external
      parse_options options
      yield self if block_given?
    end

    string_attr_accessor :display, :location, :tooltip

    serializable_attributes :display, :tooltip, :ref

    #Cell location of hyperlink on worksheet.
    # @return [String]
    attr_reader :ref

    # Sets the target for this hyperlink. Anything other than :external instructs the library to treat the location as an in-workbook reference.
    # @param [Symbol] target
    def target=(target)
      @target = target
    end

    # Sets the cell location of this hyperlink in the worksheet
    # @param [String|Cell] cell_reference The string reference or cell that defines where this hyperlink shows in the worksheet.
    def ref=(cell_reference)
      cell_reference = cell_reference.r if cell_reference.is_a?(Cell)
      Axlsx::validate_string cell_reference
      @ref = cell_reference
    end

    # The relationship instance for this hyperlink.
    # A relationship is only required if `@target` is `:external`. If not, this method will simply return `nil`.
    # @see #target=
    # @return [Relationship]
    def relationship
      return unless @target == :external
      Relationship.new(self, HYPERLINK_R, location, :target_mode => :External)
    end

    # Seralize the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str='')
      str << '<hyperlink '
      serialized_attributes str, location_or_id
      str << '/>'
    end

    # The values to be used in serialization based on the target.
    # location should only be specified for non-external targets.
    # r:id should only be specified for external targets.
    # @return [Hash]
    def location_or_id
      @target == :external ?  { :"r:id" => relationship.Id } : { :location => Axlsx::coder.encode(location) }
    end
  end
end
