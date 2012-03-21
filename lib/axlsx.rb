# encoding: UTF-8
require 'axlsx/version.rb'

require 'axlsx/util/simple_typed_list.rb'
require 'axlsx/util/constants.rb'
require 'axlsx/util/validators.rb'
require 'axlsx/util/storage.rb'
require 'axlsx/util/cbf.rb'
require 'axlsx/util/ms_off_crypto.rb'


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
require 'axlsx/table/table.rb'
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

# xlsx generation with charts, images, automated column width, customizable styles and full schema validation. Axlsx excels at helping you generate beautiful Office Open XML Spreadsheet documents without having to understand the entire ECMA specification. Check out the README for some examples of how easy it is. Best of all, you can validate your xlsx file before serialization so you know for sure that anything generated is going to load on your client's machine.
module Axlsx
  # determines the cell range for the items provided
  def self.cell_range(items)
    return "" unless items.first.is_a? Cell
    ref = "'#{items.first.row.worksheet.name}'!" +
      "#{items.first.r_abs}"
    ref += ":#{items.last.r_abs}" if items.size > 1
    ref
  end

  def self.name_to_indices(name)
    raise ArgumentError, 'invalid cell name' unless name.size > 1
    v = name[/[A-Z]+/].reverse.chars.reduce({:base=>1, :i=>0}) do  |val, c|
      val[:i] += ((c.bytes.first - 65) + val[:base]); val[:base] *= 26; val
    end

    [v[:i]-1, ((name[/[1-9][0-9]*/]).to_i)-1]

  end
end
