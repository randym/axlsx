# encoding: UTF-8
module Axlsx
  # Selection options for worksheet panes.
  #
  # @note The recommended way to manage the selection pane options is via SheetView#add_selection
  # @see SheetView#add_selection
  class Selection

    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes

    # Creates a new {Selection} object
    # @option options [Cell, String] active_cell Active Cell Location
    # @option options [Integer] active_cell_id Active Cell Index
    # @option options [Symbol] pane Pane
    # @option options [String] sqref Sequence of References
    def initialize(options={})
      @active_cell = @active_cell_id = @pane = @sqref = nil
      parse_options options
    end

    serializable_attributes :active_cell, :active_cell_id, :pane, :sqref
    # Active Cell Location
    # Location of the active cell.
    # @see type
    # @return [String]
    # default nil
    attr_reader :active_cell

    # Active Cell Index
    # 0-based index of the range reference (in the array of references listed in sqref)
    # containing the active cell. Only used when the selection in sqref is not contiguous.
    # Therefore, this value needs to be aware of the order in which the range references are
    # written in sqref.
    # When this value is out of range then activeCell can be used.
    # @see type
    # @return [Integer]
    # default nil
    attr_reader :active_cell_id

    # Pane
    # The pane to which this selection belongs.
    # Options are 
    #  * bottom_left:  Bottom left pane, when both vertical and horizontal
    #                  splits are applied. This value is also used when only
    #                  a horizontal split has been applied, dividing the pane 
    #                  into upper and lower regions. In that case, this value 
    #                  specifies the bottom pane.
    #  * bottom_right: Bottom right pane, when both vertical and horizontal
    #                  splits are applied.
    #  * top_left:     Top left pane, when both vertical and horizontal splits
    #                  are applied. This value is also used when only a horizontal 
    #                  split has been applied, dividing the pane into upper and lower
    #                  regions. In that case, this value specifies the top pane.
    #                  This value is also used when only a vertical split has
    #                  been applied, dividing the pane into right and left
    #                  regions. In that case, this value specifies the left pane
    #  * top_right:    Top right pane, when both vertical and horizontal
    #                  splits are applied. This value is also used when only
    #                  a vertical split has been applied, dividing the pane 
    #                  into right and left regions. In that case, this value 
    #                  specifies the right pane.
    # @see type
    # @return [Symbol]
    # default nil
    attr_reader :pane

    # Sequence of References
    # Range of the selection. Can be non-contiguous set of ranges.
    # @see type
    # @return [String]
    # default nil
    attr_reader :sqref

    # @see active_cell
    def active_cell=(v)
      cell = (v.class == Axlsx::Cell ? v.r_abs : v)
      Axlsx::validate_string(cell)
      @active_cell = cell
    end

    # @see active_cell_id
    def active_cell_id=(v); Axlsx::validate_unsigned_int(v); @active_cell_id = v end

    # @see pane
    def pane=(v)
      Axlsx::validate_pane_type(v)
      @pane = Axlsx::camel(v, false)
    end

    # @see sqref
    def sqref=(v); Axlsx::validate_string(v); @sqref = v end

    # Serializes the data validation
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      serialized_tag 'selection', str
    end
  end
end
