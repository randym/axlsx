module Axlsx

  # A numeric data source for use by charts.
  class NumDataSource

    include Axlsx::OptionsParser

    # creates a new NumDataSource object
    # @option options [Array] data An array of Cells or Numeric objects
    # @option options [Symbol] tag_name see tag_name
    def initialize(options={})
      # override these three in child classes
      @data_type ||= NumData
      @tag_name ||= :val
      @ref_tag_name ||= :numRef

      @f = nil
      @data = @data_type.new(options)
      if options[:data] && options[:data].first.is_a?(Cell)
        @f = Axlsx::cell_range(options[:data])
      end
      parse_options options
    end


    # The tag name to use when serializing this data source.
    # Only items defined in allowed_tag_names are allowed
    # @return [Symbol]
    attr_reader :tag_name

    attr_reader :data

    # allowed element tag names
    # @return [Array]
    def self.allowed_tag_names
      [:yVal, :val]
    end

     # sets the tag name for this data source
    # @param [Symbol] v One of the allowed_tag_names
    def tag_name=(v)
      Axlsx::RestrictionValidator.validate "#{self.class.name}.tag_name", self.class.allowed_tag_names, v
      @tag_name = v
    end

    # serialize the object
    # @param [String] str
    def to_xml_string(str="")
      str << '<c:' << tag_name.to_s << '>'
      if @f
        str << '<c:' << @ref_tag_name.to_s << '>'
        str << '<c:f>' << @f.to_s << '</c:f>'
      end
      @data.to_xml_string str
      if @f
        str << '</c:' << @ref_tag_name.to_s << '>'
      end
      str << '</c:' << tag_name.to_s << '>'
    end
  end
end

