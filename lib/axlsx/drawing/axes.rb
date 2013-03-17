module Axlsx

  class Axes

    def initialize(options={})
      options.each do |name, axis_class|
        add_axis(name, axis_class)
      end
    end

    def [](name)
      axes.assoc(name)[1]
    end

    def to_xml_string(str = '', options = {})
      if options[:ids]
        axes.inject(str) { |string, axis| string << '<c:axId val="' << axis[1].id.to_s << '"/>' }
      else
        axes.each { |axis| axis[1].to_xml_string(str) }
      end
    end

    def add_axis(name, axis_class)
      axis = axis_class.new
      set_cross_axis(axis)
      axes << [name, axis]
    end

    private

    def axes
      @axes ||= []
    end

    def set_cross_axis(axis)
      axes.first[1].cross_axis = axis if axes.size == 1
      axis.cross_axis = axes.first[1] unless axes.empty?
    end
  end
end
