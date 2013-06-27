module Axlsx
  
  # The Axes class creates and manages axis information and
  # serialization for charts.
  class Axes

    # @param [Hash] options options used to generate axis each key
    # should be an axis name like :val_axis and its value should be the
    # class of the axis type to construct. The :cat_axis, if there is one,
    # must come first (we assume a Ruby 1.9+ Hash or an OrderedHash).
    def initialize(options={})
      raise(ArgumentError, "CatAxis must come first") if options.keys.include?(:cat_axis) && options.keys.first != :cat_axis
      options.each do |name, axis_class|
        add_axis(name, axis_class)
      end
    end

    # [] provides assiciative access to a specic axis store in an axes
    # instance. 
    # @return [Axis]
    def [](name)
      axes.assoc(name)[1]
    end

    # Serializes the object
    # @param [String] str
    # @param [Hash] options
    # @option options ids
    # If the ids option is specified only the axis identifier is
    # serialized. Otherwise, each axis is serialized in full. 
    def to_xml_string(str = '', options = {})
      if options[:ids]
        # CatAxis must come first in the XML (for Microsoft Excel at least)
        sorted = axes.sort_by { |name, axis| axis.kind_of?(CatAxis) ? 0 : 1 }
        sorted.inject(str) { |string, axis| string << '<c:axId val="' << axis[1].id.to_s << '"/>' }        
      else
        axes.each { |axis| axis[1].to_xml_string(str) }
      end
    end

    # Adds an axis to the collection
    # @param [Symbol] name The name of the axis
    # @param [Axis] axis_class The axis class to generate
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
