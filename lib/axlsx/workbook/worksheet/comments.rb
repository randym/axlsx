# -*- coding: utf-8 -*-
module Axlsx

  class Comments

    # a collection of the comment authors
    # @return [SimpleTypedList]
    attr_reader :authors

    # a collection of comment objects
    # @return [SimpleTypedList]
    attr_reader :comment_list

    # the vml_drawing that holds the shapes for comments
    # @return [VmlDrawing]
    attr_reader :vml_drawing

    # The worksheet that these comments belong to
    # @return [Worksheet]
    attr_reader :worksheet

    # The index of this collection in the workbook. Effectively the index of the worksheet.
    # @return [Integer]
    def index
      @worksheet.index
    end

    # The part name for this object
    # @return [String]
    def pn
      "#{COMMENT_PN % (index+1)}"
    end

    # Creates a new Comments object
    # @param [Worksheet] worksheet The sheet that these comments belong to.
    def initialize(worksheet)
      raise ArgumentError, "you must provide a worksheet" unless worksheet.is_a?(Worksheet)
      @worksheet = worksheet
      @authors = SimpleTypedList.new String
      @comment_list = SimpleTypedList.new Comment
      @vml_drawing = VmlDrawing.new(self)
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
      yield @comment_list.last if block_given?
      @comment_list.last
    end

    # serialize the object
    # @param [String] str
    # @return [String]
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
      str << '</commentList></comments>'

    end

  end

end
