module Axlsx

  class SignatureProperty

  end

  TIME_FORMAT = "YYYY-MM-DDThh:mm:ssTZD"
  TIME_FORMAT_RUBY = "%Y-%M-%dT%H:%m:%sZ"
  attr_accessor :id
  attr_accessor :target

  attr_reader :signature_time

  def signature_time=(time)
    DataTypeValidator.validate 'SignatureProperty#signature_time', Time, time 
    @signature_time ||= SignatureTime.new(time)
  end

  def signature_info
    @signature_info ||= SignatureInfo.new
    yield @signature_info if block_given?
  end

  def to_xml_string(str = '')
    str << '<SignatureProperty Id="' << id << '" Target="' << target << '">'
    signature_time.to_xml_string(str) if signature_time
    signature_info.to_xml_string(str) if signature_info
    str << '</SignaturePropery>'
  end

end
