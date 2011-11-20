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


module Axlsx

  #required gems
  require 'Nokogiri'
  require 'active_support/core_ext/object/instance_variables'
  require 'active_support/inflector'
  require 'rmagick'
  require 'zip/zip'

  #core dependencies
  require 'bigdecimal'
  require 'time'
  require 'CGI'

  # determines the cell range for the items provided
  def self.cell_range(items)
    return "" unless items.first.is_a? Cell          
    "#{items.first.row.worksheet.name}!" +
      "#{items.first.r_abs}:#{items.last.r_abs}"
  end
  
  
end
