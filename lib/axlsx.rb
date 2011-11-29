
Encoding::default_internal = 'UTF-8' unless RUBY_VERSION < '1.9'
Encoding::default_external = 'UTF-8' unless RUBY_VERSION < '1.9'

require 'axlsx/util/simple_typed_list.rb'
require 'axlsx/util/constants.rb'
require 'axlsx/util/validators.rb'
require 'axlsx/stylesheet/styles.rb'

require 'axlsx/doc_props/app.rb'
require 'axlsx/doc_props/core.rb'
require 'axlsx/content_type/content_type.rb'
require 'axlsx/rels/relationships.rb'

require 'axlsx/drawing/drawing.rb'
require 'axlsx/workbook/workbook.rb'
require 'axlsx/package.rb'

#required gems
require 'Nokogiri'
require 'active_support/core_ext/object/instance_variables'
require 'active_support/inflector'
require 'rmagick'
require 'zip/zip'

#core dependencies
require 'bigdecimal'
require 'time'


# xlsx generation with charts, images, automated column width, customizable styles and full schema validation. Axlsx excels at helping you generate beautiful Office Open XML Spreadsheet documents without having to understand the entire ECMA specification. Check out the README for some examples of how easy it is. Best of all, you can validate your xlsx file before serialization so you know for sure that anything generated is going to load on your client's machine.
module Axlsx
  # determines the cell range for the items provided
  def self.cell_range(items)
    return "" unless items.first.is_a? Cell          
    ref = "#{items.first.row.worksheet.name}!" +
      "#{items.first.r_abs}"
    ref += ":#{items.last.r_abs}" if items.size > 1
    ref
  end
end
