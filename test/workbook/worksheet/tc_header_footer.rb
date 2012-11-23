require 'tc_helper'

class TestHeaderFooter < Test::Unit::TestCase

  def setup
    @p = Axlsx::Package.new
    ws = @p.workbook.add_worksheet :name => 'test'
    @hf = ws.header_footer
  end

  def test_initialize
    assert_equal(nil, @hf.odd_header)
    assert_equal(nil, @hf.odd_footer)

    assert_equal(nil, @hf.even_header)
    assert_equal(nil, @hf.even_footer)

    assert_equal(nil, @hf.first_header)
    assert_equal(nil, @hf.first_footer)

    assert_equal(nil, @hf.different_first)
    assert_equal(nil, @hf.different_odd_even)
  end

  def test_initialize_with_options
    header_footer = {
      :odd_header => 'oh',
      :odd_footer => 'of',

      :even_header => 'eh',
      :even_footer => 'ef',

      :first_header => 'fh',
      :first_footer => 'ff',

      :different_first => true,
      :different_odd_even => true
    }
    optioned = @p.workbook.add_worksheet(:name => 'optioned', :header_footer => header_footer).header_footer

    assert_equal('oh', optioned.odd_header)
    assert_equal('of', optioned.odd_footer)

    assert_equal('eh', optioned.even_header)
    assert_equal('ef', optioned.even_footer)

    assert_equal('fh', optioned.first_header)
    assert_equal('ff', optioned.first_footer)

    assert_equal(true, optioned.different_first)
    assert_equal(true, optioned.different_odd_even)
  end

  def test_string_attributes
    %w(odd_header odd_footer even_header even_footer first_header first_footer).each do |attr|
      assert_raise(ArgumentError, 'only strings allowed in string attributes') { @hf.send("#{attr}=", 1) }
      assert_nothing_raised { @hf.send("#{attr}=", 'test_string') }
    end
  end

  def test_boolean_attributes
   %w(different_first different_odd_even).each do |attr|
      assert_raise(ArgumentError, 'only booleanish allowed in string attributes') { @hf.send("#{attr}=", 'foo') }
      assert_nothing_raised { @hf.send("#{attr}=", 1) }
    end
  end

  def test_set_all_values
    @hf.set(
      :odd_header => 'oh',
      :odd_footer => 'of',

      :even_header => 'eh',
      :even_footer => 'ef',

      :first_header => 'fh',
      :first_footer => 'ff',

      :different_first => true,
      :different_odd_even => true
    )

    assert_equal('oh', @hf.odd_header)
    assert_equal('of', @hf.odd_footer)

    assert_equal('eh', @hf.even_header)
    assert_equal('ef', @hf.even_footer)

    assert_equal('fh', @hf.first_header)
    assert_equal('ff', @hf.first_footer)

    assert_equal(true, @hf.different_first)
    assert_equal(true, @hf.different_odd_even)
  end

  def test_to_xml_all_values
    @hf.set(
      :odd_header => 'oh',
      :odd_footer => 'of',

      :even_header => 'eh',
      :even_footer => 'ef',

      :first_header => 'fh',
      :first_footer => 'ff',

      :different_first => true,
      :different_odd_even => true
    )

    doc = Nokogiri::XML.parse(@hf.to_xml_string)
    assert_equal(1, doc.xpath(".//headerFooter[@differentFirst='true'][@differentOddEven='true']").size)

    assert_equal(1, doc.xpath(".//headerFooter/oddHeader").size)
    assert_equal('oh', doc.xpath(".//headerFooter/oddHeader").text)
    assert_equal(1, doc.xpath(".//headerFooter/oddFooter").size)
    assert_equal('of', doc.xpath(".//headerFooter/oddFooter").text)

    assert_equal(1, doc.xpath(".//headerFooter/evenHeader").size)
    assert_equal('eh', doc.xpath(".//headerFooter/evenHeader").text)
    assert_equal(1, doc.xpath(".//headerFooter/evenFooter").size)
    assert_equal('ef', doc.xpath(".//headerFooter/evenFooter").text)

    assert_equal(1, doc.xpath(".//headerFooter/firstHeader").size)
    assert_equal('fh', doc.xpath(".//headerFooter/firstHeader").text)
    assert_equal(1, doc.xpath(".//headerFooter/firstFooter").size)
    assert_equal('ff', doc.xpath(".//headerFooter/firstFooter").text)
  end

  def test_to_xml_some_values
    @hf.set(
      :odd_header => 'oh',
      :different_odd_even => false
    )

    doc = Nokogiri::XML.parse(@hf.to_xml_string)
    assert_equal(1, doc.xpath(".//headerFooter[@differentOddEven='false']").size)
    assert_equal(0, doc.xpath(".//headerFooter[@differentFirst]").size)

    assert_equal(1, doc.xpath(".//headerFooter/oddHeader").size)
    assert_equal('oh', doc.xpath(".//headerFooter/oddHeader").text)
    assert_equal(0, doc.xpath(".//headerFooter/oddFooter").size)

    assert_equal(0, doc.xpath(".//headerFooter/evenHeader").size)
    assert_equal(0, doc.xpath(".//headerFooter/evenFooter").size)

    assert_equal(0, doc.xpath(".//headerFooter/firstHeader").size)
    assert_equal(0, doc.xpath(".//headerFooter/firstFooter").size)
  end
end

