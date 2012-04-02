# encoding: UTF-8
module Axlsx
  # a hyperlink object adds an action to an image when clicked so that when the image is clicked the link is fecthed.
  # @note using the hyperlink option when calling add_image on a drawing object is the recommended way to manage hyperlinks
  # @see {file:README} README
  class Hyperlink

    # The destination of the hyperlink stored in the drawing's relationships document.
    # @return [String]
    attr_accessor :href

    # The spec says: Specifies the URL when it has been determined by the generating application that the URL is invalid. That is the generating application can still store the URL but it is known that this URL is not correct.
    #
    # What exactly that means is beyond me so if you ever use this, let me know!
    # @return [String]
    attr_accessor :invalidUrl

    #An action to take when the link is clicked. The specification says "This can be used to specify a slide to be navigated to or a script of code to be run." but in most cases you will not need to do anything with this. MS does reserve a few interesting strings. @see http://msdn.microsoft.com/en-us/library/ff532419%28v=office.12%29.aspx
    # @return [String]
    attr_accessor :action

    # Specifies if all sound events should be terminated when this link is clicked.
    # @return [Boolean]
    attr_reader :endSnd

    # @see endSnd
    # @param [Boolean] v The boolean value indicating the termination of playing sounds on click
    # @return [Boolean]
    def endSnd=(v) Axlsx::validate_boolean(v); @endSnd = v end

    # indicates that the link has already been clicked.
    # @return [Boolean]
    attr_reader :highlightClick

    # @see highlightClick
    # @param [Boolean] v The value to assign
    def highlightClick=(v) Axlsx::validate_boolean(v); @highlightClick = v end

    # From the specs: Specifies whether to add this URI to the history when navigating to it. This allows for the viewing of this presentation without the storing of history information on the viewing machine. If this attribute is omitted, then a value of 1 or true is assumed.
    # @return [Boolean]
    attr_reader :history

    # @see history
    # param [Boolean] v The value to assing
    def history=(v) Axlsx::validate_boolean(v); @history = v end

    # From the specs: Specifies the target frame that is to be used when opening this hyperlink. When the hyperlink is activated this attribute is used to determine if a new window is launched for viewing or if an existing one can be used. If this attribute is omitted, than a new window is opened.
    # @return [String]
    attr_accessor :tgtFrame

    # Text to show when you mouse over the hyperlink. If you do not set this, the href property will be shown.
    # @return [String]
    attr_accessor :tooltip

    #Creates a hyperlink object
    # parent must be a Pic for now, although I expect that other object support this tag and its cNvPr parent
    # @param [Pic] parent
    # @option options [String] tooltip message shown when hyperlinked object is hovered over with mouse.
    # @option options [String] tgtFrame Target frame for opening hyperlink
    # @option options [String] invalidUrl supposedly use to store the href when we know it is an invalid resource.
    # @option options [String] href the target resource this hyperlink links to.
    # @option options [String] action A string that can be used to perform specific actions. For excel please see this reference: http://msdn.microsoft.com/en-us/library/ff532419%28v=office.12%29.aspx
    # @option options [Boolean] endSnd terminate any sound events when processing this link
    # @option options [Boolean] history include this link in the list of visited links for the applications history.
    # @option options [Boolean] highlightClick indicate that the link has already been visited.
    def initialize(parent, options={})
      DataTypeValidator.validate "Hyperlink.parent", [Pic], parent
      @parent = parent
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
      yield self if block_given?

    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      h =  self.instance_values.merge({:'r:id' => "rId#{id}", :'xmlns:r' => XML_NS_R })
      h.delete('href')
      h.delete('parent')
      str << '<a:hlinkClick '
      str << h.map { |key, value| '' << key.to_s << '="' << value.to_s << '"' }.join(' ')
      str << '/>'
    end

    private
    # The relational ID for this hyperlink
    # @return [Integer]
    def id
      @parent.anchor.drawing.charts.size + @parent.anchor.drawing.images.size + @parent.anchor.drawing.hyperlinks.index(self) + 1
    end

  end
end
