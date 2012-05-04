module Axlsx

  class Comments

    # a collection of the comment authors
    # @return [SimpleTypedList]
    attr_reader :authors

    # a collection of comment objects
    # @return [SimpleTypedList]
    attr_reader :comment_list


    # The worksheet that these comments belong to
    # @return [Worksheet]
    attr_reader :worksheet

    # Creates a new Comments object
    # @param [Worksheet] worksheet The sheet that these comments belong to.
    def initialize(worksheet)
      raise ArgumentError, "you must provide a worksheet" unless worksheet.is_a?(Worksheet)
      @worksheet = worksheet
      @authors = SimpleTypedList.new String
      @comment_list = SimpleTypedList.new Comment
    end

    # Adds a new comment to the worksheet that owns these comments.
    # @note the author, text and ref options are required
    # @option options [String] author The name of the author for this comment
    # @option options [String] text The text for this comment
    # @option options [Stirng|Cell] ref The cell that this comment is attached to.
    def add_comment(options={})
      raise ArgumentError, "Comment require an author" unless options[:author]
      raise ArgumentError, "Comment requires text" unless options[:text]
      raise ArgumentError, "Comment requires ref" unless options[:ref]
      options[:author_index] = @authors.index(options[:author]) || @authors << options[:author]
      @comment_list << Comment.new(self, options)
      @comment_list.last
    end

    def to_xml_string(str="")
      str << '<?xml version="1.0" encoding="UTF-8"?>'
      str << '<comments xmlns="' << XML_NS << '">'
      str << '<authors>'
      authors.each do  |author|
        str << '<author>' << author.to_s << '</author>'
      end
      str << '</authors>'
      str << '<commentList>'
      comment_list.each do |comment|
        comment.to_xml_string str
      end
      str << '<commentList></comments>'

    end

  end

  class Comment

    attr_reader :text

    attr_reader :author_index

    attr_reader :comments

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

    def pn
      "#{COMMENT_PN % (index+1)}"
    end

    # The index of this comment
    # @return [Integer]
    def index
      @comments.comment_list.index(self)
    end

    def ref=(v)
      Axlsx::DataTypeValidator.validate "Comment.ref", [String, Cell], v
      @ref = v if v.is_a?(String)
      @ref = v.r if v.is_a?(Cell)
    end

    def text=(v)
      Axlsx::validate_string(v)
      @text = v
    end

    def author_index=(v)
      Axlsx::validate_unsigned_int(v)
      @author_index = v
    end

    def to_xml_string(str = "")
      str << '<comment ref="' << ref << '" authorId="' << author_index << '">'
      str << '<t xml:space="preserve">' << text << '</t>'
      str << '</comment>'
    end

  end

end
