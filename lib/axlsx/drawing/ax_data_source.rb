module Axlsx

  # An axis data source that can contain referenced or literal strings or numbers
  # @note only string data types are supported - mainly because we have not implemented a chart type that requires a numerical axis value
  class AxDataSource < NumDataSource

    # creates a new NumDataSource object
    # @option options [Array] data An array of Cells or Numeric objects
    # @option options [Symbol] tag_name see tag_name
    def initialize(options={})
      @tag_name = :cat
      @data_type = StrData
      @ref_tag_name = :strRef
      super(options)
    end

    # allowed element tag names for serialization
    # @return [Array]
    def self.allowed_tag_names
      [:xVal, :cat]
    end

  end

end

