# encoding: UTF-8
require 'htmlentities'
require 'axlsx/version.rb'

require 'axlsx/util/simple_typed_list.rb'
require 'axlsx/util/constants.rb'
require 'axlsx/util/validators.rb'
require 'axlsx/util/accessors.rb'
require 'axlsx/util/serialized_attributes'
require 'axlsx/util/options_parser'
# to be included with parsable intitites.
#require 'axlsx/util/parser.rb'

require 'axlsx/stylesheet/styles.rb'

require 'axlsx/doc_props/app.rb'
require 'axlsx/doc_props/core.rb'
require 'axlsx/content_type/content_type.rb'
require 'axlsx/rels/relationships.rb'

require 'axlsx/drawing/drawing.rb'
require 'axlsx/workbook/workbook.rb'
require 'axlsx/package.rb'
#required gems
require 'nokogiri'
require 'zip/zip'

#core dependencies
require 'bigdecimal'
require 'time'

#if object does not have this already, I am borrowing it from active_support.
# I am a very big fan of activesupports instance_values method, but do not want to require nor include the entire
# library just for this one method.
if !Object.respond_to?(:instance_values)
  Object.send :public  # patch for 1.8.7 as it uses private scope
  Object.send :define_method, :instance_values do
    Hash[instance_variables.map { |name| [name.to_s[1..-1], instance_variable_get(name)] }]
  end
end

# xlsx generation with charts, images, automated column width, customizable styles
# and full schema validation. Axlsx excels at helping you generate beautiful
# Office Open XML Spreadsheet documents without having to understand the entire
# ECMA specification. Check out the README for some examples of how easy it is.
# Best of all, you can validate your xlsx file before serialization so you know
# for sure that anything generated is going to load on your client's machine.
module Axlsx

  # determines the cell range for the items provided
  def self.cell_range(cells, absolute=true)
    return "" unless cells.first.is_a? Cell
    cells = sort_cells(cells)
    reference = "#{cells.first.reference(absolute)}:#{cells.last.reference(absolute)}"
    if absolute
      escaped_name = cells.first.row.worksheet.name.gsub "&apos;", "''"
      "'#{escaped_name}'!#{reference}"
    else
      reference
    end
  end

  # sorts the array of cells provided to start from the minimum x,y to
  # the maximum x.y#
  # @param [Array] cells
  # @return [Array]
  def self.sort_cells(cells)
    cells.sort { |x, y| [x.index, x.row.index] <=> [y.index, y.row.index] }
  end

  #global reference html entity encoding
  # @return [HtmlEntities]
  def self.coder
    @@coder ||= ::HTMLEntities.new
  end

  # returns the x, y position of a cell
  def self.name_to_indices(name)
    raise ArgumentError, 'invalid cell name' unless name.size > 1
    # capitalization?!?
    v = name[/[A-Z]+/].reverse.chars.reduce({:base=>1, :i=>0}) do  |val, c|
      val[:i] += ((c.bytes.first - 64) * val[:base]); val[:base] *= 26; val
    end
    [v[:i]-1, ((name[/[1-9][0-9]*/]).to_i)-1]
  end

  # converts the column index into alphabetical values.
  # @note This follows the standard spreadsheet convention of naming columns A to Z, followed by AA to AZ etc.
  # @return [String]
  def self.col_ref(index)
    chars = []
    while index >= 26 do
      chars << ((index % 26) + 65).chr
      index = (index / 26).to_i - 1
    end
    chars << (index + 65).chr
    chars.reverse.join
  end

  # @return [String] The alpha(column)numeric(row) reference for this sell.
  # @example Relative Cell Reference
  #   ws.rows.first.cells.first.r #=> "A1"
  def self.cell_r(c_index, r_index)
    Axlsx::col_ref(c_index).to_s << (r_index+1).to_s
  end

  # Creates an array of individual cell references based on an excel reference range.
  # @param [String] range A cell range, for example A1:D5
  # @return [Array]
  def self.range_to_a(range)
    range.match(/^(\w+?\d+)\:(\w+?\d+)$/)
    start_col, start_row = name_to_indices($1)
    end_col,   end_row   = name_to_indices($2)
    (start_row..end_row).to_a.map do |row_num|
      (start_col..end_col).to_a.map do |col_num|
        "#{col_ref(col_num)}#{row_num+1}"
      end
    end
  end

  # performs the increadible feat of changing snake_case to CamelCase
  # @param [String] s The snake case string to camelize
  # @return [String]
  def self.camel(s="", all_caps = true)
    s = s.to_s
    s = s.capitalize if all_caps
    s.gsub(/_(.)/){ $1.upcase }
  end


  # Instructs the serializer to not try to escape cell value input. 
  # This will give you a huge speed bonus, but if you content has <, > or other xml character data
  # the workbook will be invalid and excel will complain.
  def self.trust_input
    @trust_input ||= false
  end

  # @param[Boolean] trust_me A boolean value indicating if the cell value content is to be trusted
  # @return [Boolean]
  # @see Axlsx::trust_input
  def self.trust_input=(trust_me)
    @trust_input = trust_me
  end
end
