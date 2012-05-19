module Axlsx

  # A comment is the text data for a comment
  class Comment

    # The text to render
    # @return [String]
    attr_reader :text

    # The author of this comment
    # @see Comments
    # @return [String]
    attr_reader :author

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

    #
    # The index of this author in a unique sorted list of all authors in
    # the comment.
    # @return [Integer]
    def author_index
      @comments.authors.index(author)
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

    # @see author
    def author=(v)
      @author = v
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
      pos = Axlsx::name_to_indices(ref)
      @vml_shape = VmlShape.new(:row => pos[1], :column => pos[0]) do |vml|
        vml.left_column = vml.column
        vml.right_column = vml.column + 2 
        vml.top_row = vml.row
         vml.bottom_row = vml.row + 4
      end
    end

  end
end
