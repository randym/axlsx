$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"


require 'axlsx'

# First thing to do is setup our styles. OOXML Style management is unfortunately very different from CSS so we want to use the 
# add_style helper method on the workbook styles object so we dont go insane.

# I find it easier to declare a hash and then feed that in later

class FinancialReport

  def initialize(data)
    create_styles
    prepare
    insert_data data
    finalize
    package
  end

  def style_hash

    sienna = 'A0522D'
    {
      :search_results => { :sz => 10, :b => true },
      :bold_header => { :b => true },
      :grey_bg => { :bg_color => "DEDEDE" },
      :transaction__header => { :fg_color => sienna, :b => true },
      :transaction_currency => { :fb_color => sienna, :num_fmt => 5 },
      :transactin_date => { :fg_color => sienna, :format_code => 'yyyy-mm-dd' }
    }
  end

  def styles
    @styles ||= {}
  end

  def package
    @package ||= Axlsx::Package.new
  end

  # Just a place to put some defualt data for the exercise
  def self.data
    @data = ['baked',
             "American Medical Systems Holdings, Inc.",	
     "Endo Pharmaceuticals Holdings, Inc.",
     "4371",	
     Date.new,
     2757001160, 
     2519495160, 
     116,
     3842,
     "Medical devices for urology disorders"]
  end

  #  Populates an array of ['style_name'] = style_index
  def create_styles
    package.workbook.styles do |style|
      style_hash.each do |key, value|
        styles[key] = style.add_style(value)
      end
    end
  end
  def prepare
    package.workbook.add_worksheet(:name => 'All Information') do |sheet|
      sheet.add_row [nil, 'Search Results'], :style => [nil, styles['search_results']]
    end
  end


  def insert_data(data)
    package.workbook.worksheets.first do |sheet|
      sheet.add_row data, style=> [styles[:
    end
  end

  def finalize
    # package.serialize 'financial.xlsx'
  end
end
f = FinancialReport.new(FinancialReport.data)
f.package.serialize 'finance.xlsx'
