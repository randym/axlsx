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
        str << '<c r="' << Axlsx::cell_r(column_index, row_index) << '" s="' << cell.style.to_s << '" '
        return str << '/>' if cell.value.nil?
        method = (cell.type.to_s << '_type_serialization').to_sym
        self.send(method, cell, str)
        str << '</c>'
      end 


      # builds an xml text run based on this cells attributes.
      # @param [String] str The string instance this run will be concated to.
      # @return [String]
      def run_xml_string(cell, str = '')
        if cell.is_text_run?
          data = cell.instance_values.reject{|key, value| value == nil || key == 'value' || key == 'type' }
          keys = data.keys & Cell::INLINE_STYLES
          str << "<r><rPr>"
          keys.each do |key|
            case key
            when 'font_name'
              str << "<rFont val='"<< cell.font_name << "'/>"
            when 'color'
              str << data[key].to_xml_string
            else
              str << "<" << key.to_s << " val='" << data[key].to_s << "'/>"
            end
          end
          str << "</rPr>" << "<t>" << cell.value.to_s << "</t></r>"
        else
          str << "<t>" << cell.value.to_s << "</t>"
        end
        str
      end

      # serializes cells that are type iso_8601
      # @param [Cell] cell The cell that is being serialized
      # @param [String] str The string the serialized content will be appended to.
      # @return [String]
      def iso_8601_type_serialization(cell, str='')
        value_serialization 'd', cell.value, str
      end


      # serializes cells that are type date
      # @param [Cell] cell The cell that is being serialized
      # @param [String] str The string the serialized content will be appended to.
      # @return [String]
      def date_type_serialization(cell, str='')
        value_serialization false, DateTimeConverter::date_to_serial(cell.value).to_s, str
      end

      # Serializes cells that are type time
      # @param [Cell] cell The cell that is being serialized
      # @param [String] str The string the serialized content will be appended to.
      # @return [String]
      def time_type_serialization(cell, str='')
        value_serialization false, DateTimeConverter::time_to_serial(cell.value).to_s, str
      end

      # Serializes cells that are type boolean
      # @param [Cell] cell The cell that is being serialized
      # @param [String] str The string the serialized content will be appended to.
      # @return [String]
      def boolean_type_serialization(cell, str='')
        value_serialization 'b', cell.value.to_s, str
      end

      # Serializes cells that are type float
      # @param [Cell] cell The cell that is being serialized
      # @param [String] str The string the serialized content will be appended to.
      # @return [String]
      def float_type_serialization(cell, str='')
        numeric_type_serialization cell, str
      end

      # Serializes cells that are type integer
      # @param [Cell] cell The cell that is being serialized
      # @param [String] str The string the serialized content will be appended to.
      # @return [String]
      def integer_type_serialization(cell, str = '')
        numeric_type_serialization cell, str
      end


      # Serializes cells that are type formula
      # @param [Cell] cell The cell that is being serialized
      # @param [String] str The string the serialized content will be appended to.
      # @return [String]
      def formula_serialization(cell, str='')
        str << 't="str">' << '<f>' << cell.value.to_s.sub('=', '') << '</f>'
        str << '<v>' << cell.formula_value.to_s << '</v>' unless cell.formula_value.nil?
      end

      # Serializes cells that are type inline_string
      # @param [Cell] cell The cell that is being serialized
      # @param [String] str The string the serialized content will be appended to.
      # @return [String]
      def inline_string_serialization(cell, str = '')
        str << 't="inlineStr">' << '<is>'
        run_xml_string cell, str
        str << '</is>'
      end

      # Serializes cells that are type string
      # @param [Cell] cell The cell that is being serialized
      # @param [String] str The string the serialized content will be appended to.
      # @return [String]
      def string_type_serialization(cell, str='')
        if cell.is_formula?
          formula_serialization cell, str
        elsif !cell.ssti.nil?
          value_serialization 's', cell.ssti.to_s, str
        else
          inline_string_serialization cell, str
        end
      end

      private

      def numeric_type_serialization(cell, str = '')
        value_serialization 'n', cell.value.to_s, str
      end

      def value_serialization(serialization_type, serialization_value, str = '')
        str << 't="' << serialization_type << '"' if serialization_type
        str << '><v>' << serialization_value << '</v>'
      end


    end
  end
end
