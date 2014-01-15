module Axlsx
  # Pane options for a worksheet.
  #
  # @note The recommended way to manage the pane options is via SheetView#pane
  # @see SheetView#pane
  class Pane

    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes
    # Creates a new {Pane} object
    # @option options [Symbol] active_pane Active Pane
    # @option options [Symbol] state Split State
    # @option options [Cell, String] top_left_cell Top Left Visible Cell
    # @option options [Integer] x_split Horizontal Split Position
    # @option options [Integer] y_split Vertical Split Position
    def initialize(options={})
      #defaults
      @active_pane = @state = @top_left_cell = nil
      @x_split = @y_split = 0
      parse_options options
    end

    serializable_attributes :active_pane, :state, :top_left_cell, :x_split, :y_split

    # Active Pane
    # The pane that is active.
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
    attr_reader :active_pane


    # Split State
    # Indicates whether the pane has horizontal / vertical 
    # splits, and whether those splits are frozen.
    # Options are 
    #  * frozen:       Panes are frozen, but were not split being frozen. In
    #                  this state, when the panes are unfrozen again, a single
    #                  pane results, with no split. In this state, the split 
    #                  bars are not adjustable.
    #  * frozen_split: Panes are frozen and were split before being frozen. In
    #                  this state, when the panes are unfrozen again, the split
    #                  remains, but is adjustable.
    #  * split:        Panes are split, but not frozen. In this state, the split
    #                  bars are adjustable by the user.
    # @see type
    # @return [Symbol]
    # default nil
    attr_reader :state

    # Top Left Visible Cell
    # Location of the top left visible cell in the bottom 
    # right pane (when in Left-To-Right mode).
    # @see type
    # @return [String]
    # default nil
    attr_reader :top_left_cell

    # Horizontal Split Position
    # Horizontal position of the split, in 1/20th of a point; 0 (zero)
    # if none. If the pane is frozen, this value indicates the number 
    # of columns visible in the top pane.
    # @see type
    # @return [Integer]
    # default 0
    attr_reader :x_split

    # Vertical Split Position
    # Vertical position of the split, in 1/20th of a point; 0 (zero) 
    # if none. If the pane is frozen, this value indicates the number
    # of rows visible in the left pane.
    # @see type
    # @return [Integer]
    # default 0
    attr_reader :y_split

    # @see active_pane
    def active_pane=(v)
      Axlsx::validate_pane_type(v)
      @active_pane = Axlsx::camel(v.to_s, false)
    end

    # @see state
    def state=(v)
      Axlsx::validate_split_state_type(v)
      @state = Axlsx::camel(v.to_s, false)
    end

    # @see top_left_cell
    def top_left_cell=(v)
      cell = (v.class == Axlsx::Cell ? v.r_abs : v)
      Axlsx::validate_string(cell)
      @top_left_cell = cell
    end

    # @see x_split
    def x_split=(v); Axlsx::validate_unsigned_int(v); @x_split = v end

    # @see y_split
    def y_split=(v); Axlsx::validate_unsigned_int(v); @y_split = v end

    # Serializes the data validation
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      finalize
      serialized_tag 'pane', str
    end
    private

    def finalize
      if @state == 'frozen' && @top_left_cell.nil?
        row = @y_split || 0
        column = @x_split || 0
        @top_left_cell = "#{('A'..'ZZ').to_a[column]}#{row+1}"
      end
    end
  end
end
