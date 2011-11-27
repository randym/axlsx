# ActsAsAxlsx
require 'axlsx'
module Axlsx
  
  module Ar
    
    def self.included(base)
      base.send :extend, ClassMethods
    end
    
    module ClassMethods

      # we should do what ruport did and use only, exclude and methods hashes
      def acts_as_axlsx(options={})        
        include Axlsx::Ar::InstanceMethods        
        extend Axlsx::Ar::SingletonMethods
      end
    end

    module SingletonMethods

      def to_xlsx(number = :all, options = {})
        row_style = options.delete(:style)
        header_style = options.delete(:header_style) || row_style
        types = options.delete(:types)

        data = [*find(number, options)]
        data.compact!
        data.flatten!
        columns = data.first.attributes.keys
        p = Package.new
        row_style = p.workbook.styles.add_style(row_style) unless row_style.nil?
        header_style = p.workbook.styles.add_style(header_style) unless header_style.nil?
        
        p.workbook.add_worksheet(:name=>table_name.humanize) do |sheet|
          sheet.add_row columns, :style=>header_style
          data.each do |r|
            sheet.add_row r.attributes.values, :style=>row_style, :types=>types
          end
        end
        p
      end

    end

    module InstanceMethods
    end
    

  end

end
begin
require 'active_record'
ActiveRecord::Base.send :include, Axlsx::Ar
rescue Exception=>e
  puts "Running without active record extensions"
end


