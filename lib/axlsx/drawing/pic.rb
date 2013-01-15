# encoding: UTF-8
module Axlsx
  # a Pic object represents an image in your worksheet
  # Worksheet#add_image is the recommended way to manage images in your sheets
  # @see Worksheet#add_image
  class Pic

    include Axlsx::OptionsParser

    # Creates a new Pic(ture) object
    # @param [Anchor] anchor the anchor that holds this image
    # @option options [String] name
    # @option options [String] descr
    # @option options [String] image_src
    # @option options [Array] start_at
    # @option options [Intger] width
    # @option options [Intger] height
    def initialize(anchor, options={})
      @anchor = anchor
      @hyperlink = nil
      @anchor.drawing.worksheet.workbook.images << self
      parse_options options
      start_at(*options[:start_at]) if options[:start_at]
      yield self if block_given?
      @picture_locking = PictureLocking.new(options)
    end

    # allowed file extenstions
    ALLOWED_EXTENSIONS = ['gif', 'jpeg', 'png', 'jpg']

    # The name to use for this picture
    # @return [String]
    attr_reader :name

    # A description of the picture
    # @return [String]
    attr_reader :descr

    # The path to the image you want to include
    # Only local images are supported at this time and only jpg support
    # @return [String]
    attr_reader :image_src

    # The anchor for this image
    # @return [OneCellAnchor]
    attr_reader :anchor

    # The picture locking attributes for this picture
    attr_reader :picture_locking

    attr_reader :hyperlink

    # sets or updates a hyperlink for this image.
    # @param [String] v The href value for the hyper link
    # @option options @see Hyperlink#initialize All options available to the Hyperlink class apply - however href will be overridden with the v parameter value.
    def hyperlink=(v, options={})
      options[:href] = v
      if @hyperlink.is_a?(Hyperlink)
        options.each do |o|
          @hyperlink.send("#{o[0]}=", o[1]) if @hyperlink.respond_to? "#{o[0]}="
        end
      else
        @hyperlink = Hyperlink.new(self, options)
      end
      @hyperlink
    end

    def image_src=(v)
      Axlsx::validate_string(v)
      RestrictionValidator.validate 'Pic.image_src', ALLOWED_EXTENSIONS, File.extname(v).delete('.')
      raise ArgumentError, "File does not exist" unless File.exist?(v)
      @image_src = v
    end

    # @see name
    def name=(v) Axlsx::validate_string(v); @name = v; end

    # @see descr
    def descr=(v) Axlsx::validate_string(v); @descr = v; end


    # The file name of image_src without any path information
    # @return [String]
    def file_name
      File.basename(image_src) unless image_src.nil?
    end

    # returns the extension of image_src without the preceeding '.'
    # @return [String]
    def extname
      File.extname(image_src).delete('.') unless image_src.nil?
    end 
    # The index of this image in the workbooks images collections
    # @return [Index]
    def index
      @anchor.drawing.worksheet.workbook.images.index(self)
    end

    # The part name for this image used in serialization and relationship building
    # @return [String]
    def pn
      "#{IMAGE_PN % [(index+1), extname]}"
    end

    # The relational id withing the drawing's relationships
    def id
      @anchor.drawing.charts.size + @anchor.drawing.images.index(self) + 1
    end

    # Returns a relationship object for this object
    # @return Axlsx::Relationship
    def relationship
      Relationship.new(IMAGE_R, "../#{pn}")
    end

    # providing access to the anchor's width attribute
    # @param [Integer] v
    # @see OneCellAnchor.width
    def width
      return unless @anchor.is_a?(OneCellAnchor)
      @anchor.width
    end

    # @see width
    def width=(v)
      use_one_cell_anchor unless @anchor.is_a?(OneCellAnchor)
      @anchor.width = v
    end

    # providing access to update the anchor's height attribute
    # @param [Integer] v
    # @see OneCellAnchor.width
    # @note this is a noop if you are using a TwoCellAnchor
    def height
      @anchor.height
    end

    # @see height
    # @note This is a noop if you are using a TwoCellAnchor
    def height=(v)
      use_one_cell_anchor unless @anchor.is_a?(OneCellAnchor)
      @anchor.height = v
    end

    # This is a short cut method to set the start anchor position
    # If you need finer granularity in positioning use
    # graphic_frame.anchor.from.colOff / rowOff
    # @param [Integer] x The column
    # @param [Integer] y The row
    # @return [Marker]
    def start_at(x, y)
      @anchor.from.col = x
      @anchor.from.row = y
      @anchor.from
    end

    # noop if not using a two cell anchor
    # @param [Integer] x The column
    # @param [Integer] y The row
    # @return [Marker]
    def end_at(x, y)
      use_two_cell_anchor unless @anchor.is_a?(TwoCellAnchor)
      @anchor.to.col = x
      @anchor.to.row = y
      @anchor.to
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<xdr:pic>'
      str << '<xdr:nvPicPr>'
      str << '<xdr:cNvPr id="2" name="' << name.to_s << '" descr="' << descr.to_s << '">'
      @hyperlink.to_xml_string(str) if @hyperlink.is_a?(Hyperlink)
      str << '</xdr:cNvPr><xdr:cNvPicPr>'
      picture_locking.to_xml_string(str)
      str << '</xdr:cNvPicPr></xdr:nvPicPr>'
      str << '<xdr:blipFill>'
      str << '<a:blip xmlns:r ="' << XML_NS_R << '" r:embed="rId' << id.to_s <<  '"/>'
      str << '<a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr>'
      str << '<a:xfrm><a:off x="0" y="0"/><a:ext cx="2336800" cy="2161540"/></a:xfrm>'
      str << '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>'

    end

    private

    # Changes the anchor to a one cell anchor.
    def use_one_cell_anchor
      return if @anchor.is_a?(OneCellAnchor)
      swap_anchor(OneCellAnchor.new(@anchor.drawing, :from => @anchor.from))
    end

    #changes the anchor type to a two cell anchor
    def use_two_cell_anchor
      return if @anchor.is_a?(TwoCellAnchor)
      swap_anchor(TwoCellAnchor.new(@anchor.drawing)).tap do |new_anchor|
        new_anchor.from.col = @anchor.from.col
        new_anchor.from.row = @anchor.from.row
      end
    end

    # refactoring of swapping code, law of demeter be damned!
    def swap_anchor(new_anchor)
      new_anchor.drawing.anchors.delete(new_anchor)
      @anchor.drawing.anchors[@anchor.drawing.anchors.index(@anchor)] = new_anchor
      @anchor = new_anchor
    end
  end
end
