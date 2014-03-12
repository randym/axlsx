module Axlsx
  class RichTextRun
    
    include Axlsx::OptionsParser
    
    attr_reader :value
    
    INLINE_STYLES = [:font_name, :charset,
                     :family, :b, :i, :strike, :outline,
                     :shadow, :condense, :extend, :u,
                     :vertAlign, :sz, :color, :scheme].freeze
    
    def initialize(value, options={})
      self.value = value
      parse_options(options) 
    end
    
    def value=(value)
      @value = value
    end
    
    attr_accessor :cell
    
    # The inline font_name property for the cell
    # @return [String]
    attr_reader :font_name
    # @see font_name
    def font_name=(v) set_run_style :validate_string, :font_name, v; end

    # The inline charset property for the cell
    # As far as I can tell, this is pretty much ignored. However, based on the spec it should be one of the following:
    # 0 ￼ ANSI_CHARSET
    # 1 DEFAULT_CHARSET
    # 2 SYMBOL_CHARSET
    # 77 MAC_CHARSET
    # 128 SHIFTJIS_CHARSET
    # 129 ￼ HANGUL_CHARSET
    # 130 ￼ JOHAB_CHARSET
    # 134 ￼ GB2312_CHARSET
    # 136 ￼ CHINESEBIG5_CHARSET
    # 161 ￼ GREEK_CHARSET
    # 162 ￼ TURKISH_CHARSET
    # 163 ￼ VIETNAMESE_CHARSET
    # 177 ￼ HEBREW_CHARSET
    # 178 ￼ ARABIC_CHARSET
    # 186 ￼ BALTIC_CHARSET
    # 204 ￼ RUSSIAN_CHARSET
    # 222 ￼ THAI_CHARSET
    # 238 ￼ EASTEUROPE_CHARSET
    # 255 ￼ OEM_CHARSET
    # @return [String]
    attr_reader :charset
    # @see charset
    def charset=(v) set_run_style :validate_unsigned_int, :charset, v; end

    # The inline family property for the cell
    # @return [Integer]
    # 1 Roman
    # 2 Swiss
    # 3 Modern
    # 4 Script
    # 5 Decorative
    attr_reader :family
    # @see family
    def family=(v)
      set_run_style :validate_family, :family, v.to_i
    end

    # The inline bold property for the cell
    # @return [Boolean]
    attr_reader :b
    # @see b
    def b=(v) set_run_style :validate_boolean, :b, v; end

    # The inline italic property for the cell
    # @return [Boolean]
    attr_reader :i
    # @see i
    def i=(v) set_run_style :validate_boolean, :i, v; end

    # The inline strike property for the cell
    # @return [Boolean]
    attr_reader :strike
    # @see strike
    def strike=(v) set_run_style :validate_boolean, :strike, v; end

    # The inline outline property for the cell
    # @return [Boolean]
    attr_reader :outline
    # @see outline
    def outline=(v) set_run_style :validate_boolean, :outline, v; end

    # The inline shadow property for the cell
    # @return [Boolean]
    attr_reader :shadow
    # @see shadow
    def shadow=(v) set_run_style :validate_boolean, :shadow, v; end

    # The inline condense property for the cell
    # @return [Boolean]
    attr_reader :condense
    # @see condense
    def condense=(v) set_run_style :validate_boolean, :condense, v; end

    # The inline extend property for the cell
    # @return [Boolean]
    attr_reader :extend
    # @see extend
    def extend=(v) set_run_style :validate_boolean, :extend, v; end

    # The inline underline property for the cell.
    # It must be one of :none, :single, :double, :singleAccounting, :doubleAccounting, true
    # @return [Boolean]
    # @return [String]
    # @note true is for backwards compatability and is reassigned to :single
    attr_reader :u
    # @see u
    def u=(v)
      v = :single if (v == true || v == 1 || v == :true || v == 'true')
      set_run_style :validate_cell_u, :u, v
    end

    # The inline color property for the cell
    # @return [Color]
    attr_reader :color
    # @param [String] v The 8 character representation for an rgb color #FFFFFFFF"
    def color=(v)
      @color = v.is_a?(Color) ? v : Color.new(:rgb=>v)
    end

    # The inline sz property for the cell
    # @return [Inteter]
    attr_reader :sz
    # @see sz
    def sz=(v) set_run_style :validate_unsigned_int, :sz, v; end

    # The inline vertical alignment property for the cell
    # this must be one of [:baseline, :subscript, :superscript]
    # @return [Symbol]
    attr_reader :vertAlign
    # @see vertAlign
    def vertAlign=(v)
      RestrictionValidator.validate :cell_vertAlign, [:baseline, :subscript, :superscript], v
      set_run_style nil, :vertAlign, v
    end

    # The inline scheme property for the cell
    # this must be one of [:none, major, minor]
    # @return [Symbol]
    attr_reader :scheme
    # @see scheme
    def scheme=(v)
      RestrictionValidator.validate :cell_scheme, [:none, :major, :minor], v
      set_run_style nil, :scheme, v
    end

    # The Shared Strings Table index for this cell
    # @return [Integer]
    attr_reader :ssti
    
    # @return [Integer] The cellXfs item index applied to this cell.
    # @raise [ArgumentError] Invalid cellXfs id if the value provided is not within cellXfs items range.
    def style=(v)
      Axlsx::validate_unsigned_int(v)
      count = styles.cellXfs.size
      raise ArgumentError, "Invalid cellXfs id" unless v < count
      @style = v
    end

    def autowidth(widtharray)
      return if value.nil?
      if styles.cellXfs[style].alignment && styles.cellXfs[style].alignment.wrap_text
        first = true
        value.to_s.split(/\r?\n/, -1).each do |line|
          if first
            first = false
          else
            widtharray << 0
          end
          widtharray[-1] += string_width(line, font_size)
        end
      else
        widtharray[-1] += string_width(value.to_s, font_size)
      end
      widtharray
    end
    
    # Utility method for setting inline style attributes
    def set_run_style(validator, attr, value)
      return unless INLINE_STYLES.include?(attr.to_sym)
      Axlsx.send(validator, value) unless validator.nil?
      self.instance_variable_set :"@#{attr.to_s}", value
    end
    
    def to_xml_string(str = '')
      valid = RichTextRun::INLINE_STYLES
      data = Hash[self.instance_values.map{ |k, v| [k.to_sym, v] }] 
      data = data.select { |key, value| valid.include?(key) && !value.nil? }
      
      str << '<r><rPr>'
      data.keys.each do |key|
        case key
        when :font_name
          str << ('<rFont val="' << font_name << '"/>')
        when :color
          str << data[key].to_xml_string
        else
          str << ('<' << key.to_s << ' val="' << xml_value(data[key]) << '"/>')
        end
      end
      clean_value = Axlsx::trust_input ? @value.to_s : ::CGI.escapeHTML(Axlsx::sanitize(@value.to_s))
      str << ('</rPr><t>' << clean_value << '</t></r>')
    end
    
    private
    
    # Returns the width of a string according to the current style
    # This is still not perfect...
    #  - scaling is not linear as font sizes increase
    def string_width(string, font_size)
      font_scale = font_size / 10.0
      string.count(Worksheet::THIN_CHARS) * font_scale
    end
    
    # we scale the font size if bold style is applied to either the style font or
    # the cell itself. Yes, it is a bit of a hack, but it is much better than using
    # imagemagick and loading metrics for every character.
    def font_size
      return sz if sz
      font = styles.fonts[styles.cellXfs[style].fontId] || styles.fonts[0]
      (font.b || (defined?(@b) && @b)) ? (font.sz * 1.5) : font.sz
    end
    
    def style
      cell.style
    end
    
    def styles
      cell.row.worksheet.styles 
    end
    
    # Converts the value to the correct XML representation (fixes issues with
    # Numbers)
    def xml_value value
      if value == true
        1
      elsif value == false
        0
      else
        value
      end.to_s
    end
  end
end
