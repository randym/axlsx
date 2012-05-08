module Axlsx

  # A comment is the text data for a comment
  class Comment

    # The text to render
    # @return [String]
    attr_reader :text

    # The index of the the author for this comment in the owning Comments object
    # @see Comments
    # @return [Integer]
    attr_reader :author_index

    # The owning Comments object
    # @return [Comments]
    attr_reader :comments


    # The string based cell position reference (e.g. 'A1') that determines the positioning of this comment
    # @return [String]
    attr_reader :ref

    # TODO
    # r (Rich Text Run)
    # rPh (Phonetic Text Run)
    # phoneticPr (Phonetic Properties)

    def initialize(comments, options={})
      raise ArgumentError, "A comment needs a parent comments object" unless comments.is_a?(Comments)
      @comments = comments
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
      yield self if block_given?
    end

    # The vml shape that will render this comment
    # @return [VmlShape]
    def vml_shape
      @vml_shape ||= initialize_vml_shape
    end

    # The index of this comment
    # @return [Integer]
    def index
      @comments.comment_list.index(self)
    end

    # @see ref
    def ref=(v)
      Axlsx::DataTypeValidator.validate "Comment.ref", [String, Cell], v
      @ref = v if v.is_a?(String)
      @ref = v.r if v.is_a?(Cell)
    end

    # @see text
    def text=(v)
      Axlsx::validate_string(v)
      @text = v
    end

    # @see author_index
    def author_index=(v)
      Axlsx::validate_unsigned_int(v)
      @author_index = v
    end

    # serialize the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = "")
      author = @comments.authors[author_index]
      str << '<comment ref="' << ref << '" authorId="' << author_index.to_s << '">'
      str << '<text><r>'
      str << '<rPr> <b/><color indexed="81"/></rPr>'
      str << '<t>' << author.to_s << ':
</t></r>'
      str << '<r>'
      str << '<rPr><color indexed="81"/></rPr>'
      str << '<t>' << text << '</t></r></text>'
      str << '</comment>'
    end

    private

    # initialize the vml shape based on this comment's ref/position in the worksheet.
    # by default, all columns are 5 columns wide and 5 rows high
    def initialize_vml_shape
      ws = self.comments.worksheet
      @vml_shape = VmlShape.new(self, :row => ws[ref].row.index, :column => ws[ref].index) do |vml|
        vml.left_column = vml.row + 1
        vml.right_column = vml.column + 4
        vml.top_row = vml.row
        vml.bottom_row = vml.row + 4
      end
    end

  end
end
