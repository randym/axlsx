$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../../lib"
require 'axlsx'

package = Axlsx::Package.new
package.workbook.add_worksheet(:name => "Basic Worksheet") do |sheet|
  sheet.add_row ["First Column", "Second", "Third"]
  sheet.add_row [1, 2, 3]
end
package.serialize 'basics.xlsx'


