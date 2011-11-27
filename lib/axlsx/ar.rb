# ActsAsAxlsx
require 'axlsx'
module Axlsx
  # Mixing module for adding acts_as_axlsx to active record base
  module Ar
    
    # Extents active record with this ojbects class method acts_as_axlsx
    def self.included(base)
      base.send :extend, ClassMethods
    end
    
    # Class methods for the mixin
    module ClassMethods

      # adds in the instance and singleton methods
      def acts_as_axlsx(options={})        
        include Axlsx::Ar::InstanceMethods        
        extend Axlsx::Ar::SingletonMethods
      end
    end

    # Singleton methods for the mixin
    module SingletonMethods

      # Maps the AR class to an Axlsx package
      # options are passed into AR find
      # @param [Symbol, :all, etc.
      # @option options [Integer] header_style to apply to the first row of field names
      # @option options [Array, Symbol] an array of Axlsx types for each cell in data rows or a single type that will be applied to all types.
      # @option options [Integer, Array] style The style to pass to Worksheet#add_row
      # @see Worksheet#add_row
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

    # Empty module - I really like ruports way of allowing :include, :only, :exclude 
    # and am looking to add something like that in the next release
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


