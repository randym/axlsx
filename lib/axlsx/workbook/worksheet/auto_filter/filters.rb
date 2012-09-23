module Axlsx

  class Filters

    def initialize(options={})
      options[:filter_items].each do |filter|
        @filters << Filter.new(filter)
      end
      options[:date_group_items].each do |date_group|
        @date_group_items << DateGroupItem.new(date_group)
      end
    end
  end
end
