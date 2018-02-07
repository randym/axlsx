module Axlsx

  # The Cell Serializer class contains the logic for serializing cells based on their type.
  class CellSerializer
    class << self
      # Calls the proper serialization method based on type.
      # @param [Integer] row_index The index of the cell's row
      # @param [Integer] column_index The index of the cell's column
      # @param [String] str The string to apend serialization to.
      # @return [String]
      def to_xml_string(row_index, column_index, cell, str='')
        str << ('<c r="' << Axlsx::cell_r(column_index, row_index) << '" s="' << cell.style.to_s << '" ')
        return str << '/>' if cell.value.nil?
        method = cell.type
        self.send(method, cell, str)
        str << '</c>'
      end

      # builds an xml text run based on this cells attributes.
      # @param [String] str The string instance this run will be concated to.
      # @return [String]
      def run_xml_string(cell, str = '')
        if cell.is_text_run?
          valid = RichTextRun::INLINE_STYLES - [:value, :type]
          data = Hash[cell.instance_values.map{ |k, v| [k.to_sym, v] }]
          data = data.select { |key, value| valid.include?(key) && !value.nil? }
          RichText.new(cell.value.to_s, data).to_xml_string(str)
        elsif cell.contains_rich_text?
          cell.value.to_xml_string(str)
        else
          str << ('<t>' << cell.clean_value << '</t>')
        end
        str
      end

      # serializes cells that are type iso_8601
      # @param [Cell] cell The cell that is being serialized
      # @param [String] str The string the serialized content will be appended to.
      # @return [String]
      def iso_8601(cell, str='')
        value_serialization 'd', cell.value, str
      end

      # serializes cells that are type date
      # @param [Cell] cell The cell that is being serialized
      # @param [String] str The string the serialized content will be appended to.
      # @return [String]
      def date(cell, str='')
        value_serialization false, DateTimeConverter::date_to_serial(cell.value).to_s, str
      end

      # Serializes cells that are type time
      # @param [Cell] cell The cell that is being serialized
      # @param [String] str The string the serialized content will be appended to.
      # @return [String]
      def time(cell, str='')
        value_serialization false, DateTimeConverter::time_to_serial(cell.value).to_s, str
      end

      # Serializes cells that are type boolean
      # @param [Cell] cell The cell that is being serialized
      # @param [String] str The string the serialized content will be appended to.
      # @return [String]
      def boolean(cell, str='')
        value_serialization 'b', cell.value.to_s, str
      end

      # Serializes cells that are type float
      # @param [Cell] cell The cell that is being serialized
      # @param [String] str The string the serialized content will be appended to.
      # @return [String]
      def float(cell, str='')
        numeric cell, str
      end

      # Serializes cells that are type integer
      # @param [Cell] cell The cell that is being serialized
      # @param [String] str The string the serialized content will be appended to.
      # @return [String]
      def integer(cell, str = '')
        numeric cell, str
      end

      # Serializes cells that are type formula
      # @param [Cell] cell The cell that is being serialized
      # @param [String] str The string the serialized content will be appended to.
      # @return [String]
      def formula_serialization(cell, str='')
        str << ('t="str"><f>' << cell.clean_value.to_s.sub('=', '') << '</f>')
        str << ('<v>' << cell.formula_value.to_s << '</v>') unless cell.formula_value.nil?
      end

      # Serializes cells that are type array formula
      # @param [Cell] cell The cell that is being serialized
      # @param [String] str The string the serialized content will be appended to.
      # @return [String]
      def array_formula_serialization(cell, str='')
        str << ('t="str">' << '<f t="array" ref="' << cell.r << '">' << cell.clean_value.to_s.sub('{=', '').sub(/}$/, '') << '</f>')
        str << ('<v>' << cell.formula_value.to_s << '</v>') unless cell.formula_value.nil?
      end

      # Serializes cells that are type inline_string
      # @param [Cell] cell The cell that is being serialized
      # @param [String] str The string the serialized content will be appended to.
      # @return [String]
      def inline_string_serialization(cell, str = '')
        str << 't="inlineStr"><is>'
        run_xml_string cell, str
        str << '</is>'
      end

      # Serializes cells that are type string
      # @param [Cell] cell The cell that is being serialized
      # @param [String] str The string the serialized content will be appended to.
      # @return [String]
      def string(cell, str='')
        if cell.is_array_formula?
          array_formula_serialization cell, str
        elsif cell.is_formula?
          formula_serialization cell, str
        elsif !cell.ssti.nil?
          value_serialization 's', cell.ssti, str
        else
          inline_string_serialization cell, str
        end
      end

      # Serializes cells that are of the type richtext
      # @param [Cell] cell The cell that is being serialized
      # @param [String] str The string the serialized content will be appended to.
      # @return [String]
      def richtext(cell, str)
        if cell.ssti.nil?
          inline_string_serialization cell, str
        else
          value_serialization 's', cell.ssti, str
        end
      end

      # Serializes cells that are of the type text
      # @param [Cell] cell The cell that is being serialized
      # @param [String] str The string the serialized content will be appended to.
      # @return [String]
      def text(cell, str)
        if cell.ssti.nil?
          inline_string_serialization cell, str
        else
          value_serialization 's', cell.ssti, str
        end
      end

      private

      def numeric(cell, str = '')
        value_serialization 'n', cell.value, str
      end

      def value_serialization(serialization_type, serialization_value, str = '')
        str << ('t="' << serialization_type.to_s << '"') if serialization_type
        str << ('><v>' << serialization_value.to_s << '</v>')
      end
    end
  end
end
