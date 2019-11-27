# -*- coding: utf-8 -*-
# frozen_string_literal: true
module Axlsx

  # Comments is a collection of Comment objects for a worksheet
  class Comments < SimpleTypedList

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
      super(Comment)
      @worksheet = worksheet
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
      self << Comment.new(self, options)
      yield last if block_given?
      last
    end

    # A sorted list of the unique authors in the contained comments
    # @return [Array]
    def authors
      map { |comment| comment.author.to_s }.uniq.sort
    end

    # The relationships required by this object
    # @return [Array]
    def relationships
      [Relationship.new(self, VML_DRAWING_R, "../#{vml_drawing.pn}"),
       Relationship.new(self, COMMENT_R, "../#{pn}")]
    end

    # serialize the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = String.new)
      str << '<?xml version="1.0" encoding="UTF-8"?>'
      str << "<comments xmlns=\"#{XML_NS}\"><authors>"
      authors.each do  |author|
        str << "<author>#{author}</author>"
      end
      str << '</authors><commentList>'
      each do |comment|
        comment.to_xml_string str
      end
      str << '</commentList></comments>'

    end

  end

end
