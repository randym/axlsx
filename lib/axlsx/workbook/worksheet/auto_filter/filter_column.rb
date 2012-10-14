module Axlsx
  # The filterColumn collection identifies a particular column in the AutoFilter
  # range and specifies filter information that has been applied to this column.
  # If a column in the AutoFilter range has no criteria specified,
  # then there is no corresponding filterColumn collection expressed for that column.
  class FilterColumn

    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes

    # Creates a new FilterColumn object
    # @note This class yeilds its filter object as that is where the vast majority of processing will be done
    # @param [Integer|Cell] col_id The zero based index for the column to which this filter will be applied
    # @param [Symbol] filter_type The symbolized class name of the filter to apply to this column.
    # @param [Hash] options options for this object and the filter
    # @option [Boolean] hidden_button @see hidden_button
    # @option [Boolean] show_button @see show_button
    def initialize(col_id, filter_type, options = {})
      RestrictionValidator.validate 'FilterColumn.filter', FILTERS, filter_type
      #Axlsx::validate_unsigned_int(col_id)
      self.col_id = col_id
      parse_options options
      @filter = Axlsx.const_get(Axlsx.camel(filter_type)).new(options)
      yield @filter if block_given?
    end

    serializable_attributes :col_id, :hidden_button, :show_button

    # Allowed filters
    FILTERS =  [:filters] #, :top10, :custom_filters, :dynamic_filters, :color_filters, :icon_filters]

    # Zero-based index indicating the AutoFilter column to which this filter information applies.
    # @return [Integer]
    attr_reader :col_id

    # The actual filter being dealt with here
    # This could be any one of the allowed filter types
    attr_reader :filter

    # Flag indicating whether the filter button is visible.
    # When the cell containing the filter button is merged with another cell,
    # the filter button can be hidden, and not drawn.
    # @return [Boolean]
    def show_button
      @show_button ||= true
    end

    # Flag indicating whether the AutoFilter button for this column is hidden.
    # @return [Boolean]
    def hidden_button
      @hidden_button ||= false
    end

    # Sets the col_id attribute for this filter column.
    # @param [Integer | Cell] column_index The zero based index of the column to which this filter applies. 
    #                         When you specify a cell, the column index will be read off the cell
    # @return [Integer]
    def col_id=(column_index)
      column_index = column_index.col if column_index.is_a?(Cell)
      Axlsx.validate_unsigned_int column_index
      @col_id = column_index
    end

    # Apply the filters for this column
    # @param [Array] row A row from a worksheet that needs to be
    # filtered.
    def apply(row, offset)
      row.hidden = @filter.apply(row.cells[offset+col_id.to_i])
    end
    # @param [Boolean] hidden Flag indicating whether the AutoFilter button for this column is hidden.
    # @return [Boolean]
    def hidden_button=(hidden)
      Axlsx.validate_boolean hidden
      @hidden_button = hidden
    end

    # Flag indicating whether the AutoFilter button is show. This is
    # undocumented in the spec, but exists in the schema file as an
    # optional attribute.
    # @param [Boolean] show Show or hide the button
    # @return [Boolean]
    def show_button=(show)
      Axlsx.validate_boolean show
      @show_botton = show
    end

    # Serialize the object to xml
    def to_xml_string(str='')
      str << "<filterColumn #{serialized_attributes}>"
      @filter.to_xml_string(str)
      str << "</filterColumn>"
    end
  end
end
