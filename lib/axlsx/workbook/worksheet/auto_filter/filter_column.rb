module Axlsx
  # The filterColumn collection identifies a particular column in the AutoFilter
  # range and specifies filter information that has been applied to this column.
  # If a column in the AutoFilter range has no criteria specified,
  # then there is no corresponding filterColumn collection expressed for that column.
  class FilterColumn

    FILTERS =  [:filters, :top10, :custom_filters, :dynamic_filters, :color_filters, :icon_filters]
    # Zero-based index indicating the AutoFilter column to which this filter information applies.
    # @return [Integer]
    attr_reader :col_id

    # Flag indicating whether the AutoFilter button for this column is hidden.
    # @return [Boolean]
    attr_reader :hidden_button

    # Flag indicating whether the filter button is visible.
    # When the cell containing the filter button is merged with another cell,
    # the filter button can be hidden, and not drawn.
    # @return [Boolean]
    attr_reader :show_button

    # Filter information applied to the column
    # @return [Filters|Top10|CustomFilters|dynamicFilters|colorFilters|iconFilters] The filter
    #         information for this filter columns.
    attr_reader :filter

    # Creates a new FilterColumn object
    # @note This class yeilds its filter object as that is where the vast majority of processing will be done
    # @param [Integer|Cell] col_id The zero based index for the column to which this filter will be applied
    # @param [Symbol] filter_type The symbolized class name of the filter to apply to this column.
    # @param [Hash] options options for this object and the filter
    # @option [Boolean] hidden_button @see hidden_button
    # @option [Boolean] show_button @see show_button

    def initialize(col_id, filter_type, options = {})
      RestrictionValidator.validate 'FilterColumn.filter', FILTERS, filter_type
      self.col_id = col_id
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
      @filter = Axlsx.const_get(Axlsx.camel(filter_type))
      yield @filter if block_given?
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

    # Sets the button_hidden attribute for this filter column 
    # @param [Boolean] hidden Flag indicating whether the AutoFilter button for this column is hidden.
    # @return [Boolean]
    def button_hidden=(hidden)
      DataValidater.validate_boolean hidden
      @button_hidden = hidden
    end

    # Serialize the object to xml
    def to_xml_string(str='')

    end

  end
end
