# encoding: UTF-8
module Axlsx
  # The picture locking class defines the locking properties for pictures in your workbook.
  class PictureLocking

    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes
    include Axlsx::Accessors

    boolean_attr_accessor :noGrp, :noSelect, :noRot, :noChangeAspect,
                            :noMove, :noResize, :noEditPoints, :noAdjustHandles,
                            :noChangeArrowheads, :noChangeShapeType

    serializable_attributes :noGrp, :noSelect, :noRot, :noChangeAspect,
                            :noMove, :noResize, :noEditPoints, :noAdjustHandles,
                            :noChangeArrowheads, :noChangeShapeType

    # Creates a new PictureLocking object
    # @option options [Boolean] noGrp
    # @option options [Boolean] noSelect
    # @option options [Boolean] noRot
    # @option options [Boolean] noChangeAspect
    # @option options [Boolean] noMove
    # @option options [Boolean] noResize
    # @option options [Boolean] noEditPoints
    # @option options [Boolean] noAdjustHandles
    # @option options [Boolean] noChangeArrowheads
    # @option options [Boolean] noChangeShapeType
    def initialize(options={})
      @noChangeAspect = true
      parse_options options
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      serialized_tag('a:picLocks', str)
    end

  end
end
