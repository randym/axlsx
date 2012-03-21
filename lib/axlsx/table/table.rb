# encoding: UTF-8
module Axlsx
  # Table
  # @note Worksheet#add_table is the recommended way to create charts for your worksheets.
  # @see README for examples
  class Table


    # The reference to the table data
    # @return [String]
    attr_reader :ref

    # The name of the table.
    # @return [String]
    attr_reader :name

    # The style for the table. 
    # @return [TableStyle]
    attr_reader :style
   
    # Creates a new Table object
    # @param [String] ref The reference to the table data.
    # @param [Sheet] ref The sheet containing the table data.
    # @option options [Cell, String] name
    # @option options [TableStyle] style
    def initialize(ref, sheet, options={})
      @ref = ref
      @sheet = sheet
      @style = nil
      @sheet.workbook.tables << self
      @name = "Table#{index+1}"
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
      yield self if block_given?
    end

    # The index of this chart in the workbooks charts collection
    # @return [Integer]
    def index
      @sheet.workbook.tables.index(self)
    end

    # The part name for this table
    # @return [String]
    def pn
      "#{TABLE_PN % (index+1)}"
    end
    
    # The relation reference id for this table
    # @return [String]
    def rId
      "rId#{index+1}"
    end

    # The name of the Table.
    # @param [String, Cell] v
    # @return [Title]
    def name=(v) 
      DataTypeValidator.validate "#{self.class}.name", [String], v
      if v.is_a?(String)
        @name = v
      end
    end

    
    # The style for the table.
    # TODO 
    # def style=(v) DataTypeValidator.validate "Chart.style", Integer, v, lambda { |arg| arg >= 1 && arg <= 48 }; @style = v; end

    # Table Serialization
    # serializes the table
    def to_xml
      builder = Nokogiri::XML::Builder.new(:encoding => ENCODING) do |xml|
        xml.table(:xmlns => XML_NS, :id => index+1, :name => @name, :displayName => @name.gsub(/\s/,'_'), :ref => @ref, :totalsRowShown => 0) {
          xml.autoFilter :ref=>@ref 
          xml.tableColumns(:count => header_cells.length) {
            header_cells.each_with_index do |cell,index|
              xml.tableColumn :id => index+1, :name => cell.value
            end
          }
          xml.tableStyleInfo :showFirstColumn=>"0", :showLastColumn=>"0", :showRowStripes=>"1", :showColumnStripes=>"0", :name=>"TableStyleMedium9"
          #TODO implement tableStyleInfo
        }
      end
      builder.to_xml(:save_with => 0)
    end

    private
    
    # get the header cells (hackish)
    def header_cells 
      header = @ref.gsub(/^(\w+)(\d+)\:(\w+)\d+$/, '\1\2:\3\2')
      @sheet[header]
    end
  end
end
