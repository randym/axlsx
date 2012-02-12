# encoding: UTF-8
module Axlsx

  # The Shared String Table class is responsible for managing and serializing common strings in a workbook.
  # While the ECMA-376 spec allows for both inline and shared strings it seems that at least some applications like Numbers (Mac)
  # and Google Docs require that the shared string table is populated in order to interoperate properly. 
  # As a developer, you should never need to directly work against this class. Simply set 'use_shared_strings'
  # on the package or workbook to generate a package that uses the shared strings table instead of inline strings.
  # @note Serialization performance is affected by using this serialization method so if you do not need interoperability
  # it is recomended that you use the default inline string method of serialization.
  class SharedStringsTable

    # The total number of strings in the workbook including duplicates
    # Empty cells are treated as blank strings
    # @return [Integer]
    attr_reader :count

    # The total number of unique strings in the workbook. 
    # @return [Integer]
    def unique_count
      @unique_cells.size
    end

    # An array of unique cells. Multiple attributes of the cell are used in comparison
    # each of these unique cells is parsed into the shared string table.     
    # @see Cell#sharable
    attr_reader :unique_cells

    # Creates a new Shared Strings Table agains an array of cells
    # @param [Array] cells This is an array of all of the cells in the workbook
    def initialize(cells)
      cells = cells.flatten.reject { |c| c.type != :string }
      @count = cells.size
      @unique_cells = []
      resolve cells
    end


    # Generate the xml document for the Shared Strings Table
    # @return [String]
    def to_xml
      builder = Nokogiri::XML::Builder.new(:encoding => ENCODING) do |xml|
        xml.sst(:xmlns => Axlsx::XML_NS, :count => count, :uniqueCount => unique_count) {
          @unique_cells.each do |cell|
            xml.si { xml.t cell.value.to_s }
          end
        }
      end
      builder.to_xml(:save_with => 0)
    end
  end

  private

  # Interate over all of the cells in the array. 
  # if our unique cells array does not contain a sharable cell, 
  # add the cell to our unique cells array and set the ssti attribute on the index of this cell in the shared strings table
  # if a sharable cell already exists in our unique_cells array, set the ssti attribute of the cell and move on.
  # @return [Array] unique cells
  def resolve(cells)
    cells.each do |cell|
      index = @unique_cells.index { |item| item.shareable(cell) }
      if index == nil
        cell.send :ssti=, @unique_cells.size
        @unique_cells << cell
      else
        cell.send :ssti=, index
      end
    end
  end
  
end
