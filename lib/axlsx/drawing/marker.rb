# encoding: UTF-8
module Axlsx
  # The Marker class defines a point in the worksheet that drawing anchors attach to.
  # @note The recommended way to manage markers is Worksheet#add_chart Markers are created for a two cell anchor based on the :start and :end options.
  # @see Worksheet#add_chart
  class Marker

    include Axlsx::OptionsParser

    # Creates a new Marker object
    # @option options [Integer] col
    # @option options [Integer] colOff
    # @option options [Integer] row
    # @option options [Integer] rowOff
    def initialize(options={})
      @col, @colOff, @row, @rowOff = 0, 0, 0, 0
      parse_options options
    end

    # The column this marker anchors to
    # @return [Integer]
    attr_reader :col

    # The offset distance from this marker's column
    # @return [Integer]
    attr_reader :colOff

    # The row this marker anchors to
    # @return [Integer]
    attr_reader :row

    # The offset distance from this marker's row
    # @return [Integer]
    attr_reader :rowOff

     # @see col
    def col=(v) Axlsx::validate_unsigned_int v; @col = v end
    # @see colOff
    def colOff=(v) Axlsx::validate_int v; @colOff = v end
    # @see row
    def row=(v) Axlsx::validate_unsigned_int v; @row = v end
    # @see rowOff
    def rowOff=(v) Axlsx::validate_int v; @rowOff = v end

    # shortcut to set the column, row position for this marker
    # @param col the column for the marker, a Cell object or a string reference like "B7"
    # or an Array.
    # @param row the row of the marker. This is ignored if the col parameter is a Cell or
    # String or Array.
    def coord(col, row=0)
      coordinates = parse_coord_args(col, row)
      self.col = coordinates[0]
      self.row = coordinates[1]
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      [:col, :colOff, :row, :rowOff].each do |k|
        str << ('<xdr:' << k.to_s << '>' << self.send(k).to_s << '</xdr:' << k.to_s << '>')
      end
    end
    private

    # handles multiple inputs for setting the position of a marker
    # @see Chart#start_at
    def parse_coord_args(x, y=0)
      if x.is_a?(String)
        x, y = *Axlsx::name_to_indices(x)
      end
      if x.is_a?(Cell)
        x, y = *x.pos
      end
      if x.is_a?(Array)
        x, y = *x
      end
      [x, y]
    end


  end

end
