module Axlsx

  # The Break class stores the details for row and column page breaks.
  # @see RowBreaks, ColBreaks
  class Break

    include Axlsx::OptionsParser
    include Axlsx::Accessors
    include Axlsx::SerializedAttributes

    # Creates a new Break object
    # @param options A hash of attribute options for this break.
    # @option options [Integer] id Zero-based row or column Id of the page break. Breaks occur above the specified row and left of the specified column.
    # @option options [Integer] min Zero-based index of start row or column of the break. For row breaks, specifies column index; for column breaks, specifies row index.
    # @option options [Integer] max Zero-based index of end row or column of the break. For row breaks, specifies column index; for column breaks, specifies row index.
    # @option options [Boolean] man Manual Break flag. 1 means the break is a manually inserted break.
    # @option option [Boolean] pt Flag indicating that a PivotTable created this break.
    def initialize(options={})
      parse_options options
      yield self if block_given?
    end

    unsigned_int_attr_accessor :id, :min, :max

    boolean_attr_accessor :man, :pt

    serializable_attributes :id, :min, :max, :man, :pt

    # serializes the break to xml
    def to_xml_string(str='')
      serialized_tag('brk', str)
    end
  end
end

