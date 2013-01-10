module Axlsx
  class CellSerializer
    class << self
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


      def iso_8601_type_serialization(cell, str='')
        value_serialization 'd', cell.value, str
      end

      def date_type_serialization(cell, str='')
        value_serialization false, DateTimeConverter::date_to_serial(cell.value).to_s, str
      end

      def time_type_serialization(cell, str='')
        value_serialization false, DateTimeConverter::time_to_serial(cell.value).to_s, str
      end

      def boolean_type_serialization(cell, str='')
        value_serialization 'b', cell.value.to_s, str
      end

      def float_type_serialization(cell, str='')
        numeric_type_serialization cell, str
      end

      def integer_type_serialization(cell, str = '')
        numeric_type_serialization cell, str
      end

      def numeric_type_serialization(cell, str = '')
        value_serialization 'n', cell.value.to_s, str
      end

      def value_serialization(serialization_type, serialization_value, str = '')
        str << 't="' << serialization_type << '"' if serialization_type
        str << '><v>' << serialization_value << '</v>'
      end 

      def formula_serialization(cell, str='')
        str << 't="str">' << '<f>' << cell.value.to_s.sub('=', '') << '</f>'
        str << '<v>' << cell.formula_value.to_s << '</v>' unless cell.formula_value.nil?
      end

      def inline_string_serialization(cell, str = '')
        str << 't="inlineStr">' << '<is>'
        run_xml_string cell, str
        str << '</is>'
      end

      def string_type_serialization(cell, str='')
        if cell.is_formula?
          formula_serialization cell, str
        elsif !cell.ssti.nil?
          value_serialization 's', cell.ssti.to_s, str
        else
          inline_string_serialization cell, str
        end
      end
    end
  end
end
