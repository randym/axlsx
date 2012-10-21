module Axlsx

  #A collection of Cfvo objects that initializes with the required
  #first two items
  class Cfvos < SimpleTypedList

    def initialize
      super(Cfvo)
      @list << Cfvo.new(:type => :min, :val => 0)
      @list << Cfvo.new(:type => :max, :val => 0)
      lock
    end

    def to_xml_string(str='')
      @list.each { |cfvo| cfvo.to_xml_string(str) }
    end
  end
end
