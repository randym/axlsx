# encoding: UTF-8
module Axlsx

  # App represents the app.xml document. The attributes for this object are primarily managed by the application the end user uses to edit the document. None of the attributes are required to serialize a valid xlsx object.
  # @see shared-documentPropertiesExtended.xsd
  # @note Support is not implemented for the following complex types:
  #
  #    HeadingPairs (VectorVariant),
  #    TitlesOfParts (VectorLpstr),
  #    HLinks (VectorVariant),
  #    DigSig (DigSigBlob)
  class App

    include Axlsx::OptionsParser

    # Creates an App object
    # @option options [String] template
    # @option options [String] manager
    # @option options [Integer] pages
    # @option options [Integer] words
    # @option options [Integer] characters
    # @option options [String] presentation_format
    # @option options [Integer] lines
    # @option options [Integer] paragraphs
    # @option options [Integer] slides
    # @option options [Integer] notes
    # @option options [Integer] total_time
    # @option options [Integer] hidden_slides
    # @option options [Integer] m_m_clips
    # @option options [Boolean] scale_crop
    # @option options [Boolean] links_up_to_date
    # @option options [Integer] characters_with_spaces
    # @option options [Boolean] share_doc
    # @option options [String] hyperlink_base
    # @option options [String] hyperlinks_changed
    # @option options [String] application
    # @option options [String] app_version
    # @option options [Integer] doc_security
    def initialize(options={})
      parse_options options
    end

    # @return [String] The name of the document template.
    attr_reader :template
    alias :Template :template

    # @return [String] The name of the manager for the document.
    attr_reader :manager
    alias :Manager :manager

    # @return [String] The name of the company generating the document.
    attr_reader :company
    alias :Company :company

    # @return [Integer] The number of pages in the document.
    attr_reader :pages
    alias :Pages :pages

    # @return [Integer] The number of words in the document.
    attr_reader :words
    alias :Words :words

    # @return [Integer] The number of characters in the document.
    attr_reader :characters
    alias :Characters :characters

    # @return [String] The intended format of the presentation.
    attr_reader :presentation_format
    alias :PresentationFormat :presentation_format

    # @return [Integer] The number of lines in the document.
    attr_reader :lines
    alias :Lines :lines

    # @return [Integer] The number of paragraphs in the document
    attr_reader :paragraphs
    alias :Paragraphs :paragraphs

    # @return [Intger] The number of slides in the document.
    attr_reader :slides
    alias :Slides :slides

    # @return [Integer] The number of slides that have notes.
    attr_reader :notes
    alias :Notes :notes

    # @return [Integer] The total amount of time spent editing.
    attr_reader :total_time
    alias :TotalTime :total_time

    # @return [Integer] The number of hidden slides.
    attr_reader :hidden_slides
    alias :HiddenSlides :hidden_slides

    # @return [Integer] The total number multimedia clips
    attr_reader :m_m_clips
    alias :MMClips :m_m_clips

    # @return [Boolean] The display mode for the document thumbnail.
    attr_reader :scale_crop
    alias :ScaleCrop :scale_crop

    # @return [Boolean] The links in the document are up to date.
    attr_reader :links_up_to_date
    alias :LinksUpToDate :links_up_to_date

    # @return [Integer] The number of characters in the document including spaces.
    attr_reader :characters_with_spaces
    alias :CharactersWithSpaces :characters_with_spaces

    # @return [Boolean] Indicates if the document is shared.
    attr_reader :shared_doc
    alias :SharedDoc :shared_doc

    # @return [String] The base for hyper links in the document.
    attr_reader :hyperlink_base
    alias :HyperlinkBase :hyperlink_base

    # @return [Boolean] Indicates that the hyper links in the document have been changed.
    attr_reader :hyperlinks_changed
    alias :HyperlinksChanged :hyperlinks_changed

    # @return [String] The name of the application
    attr_reader :application
    alias :Applicatoin :application

    # @return [String] The version of the application.
    attr_reader :app_version
    alias :AppVersion :app_version

    # @return [Integer] Document security
    attr_reader :doc_security
    alias :DocSecurity :doc_security

    # Sets the template property of your app.xml file
    def template=(v) Axlsx::validate_string v; @template = v; end
    alias :Template= :template=

    # Sets the manager property of your app.xml file
    def manager=(v) Axlsx::validate_string v; @manager = v; end
    alias :Manager= :manager=

    # Sets the company property of your app.xml file
    def company=(v) Axlsx::validate_string v; @company = v; end
    alias :Company= :company=
    # Sets the pages property of your app.xml file
    def pages=(v) Axlsx::validate_int v; @pages = v; end

    # Sets the words property of your app.xml file
    def words=(v) Axlsx::validate_int v; @words = v; end
    alias :Words= :words=

    # Sets the characters property of your app.xml file
    def characters=(v) Axlsx::validate_int v; @characters = v; end
    alias :Characters= :characters=

    # Sets the presentation_format property of your app.xml file
    def presentation_format=(v) Axlsx::validate_string v; @presentation_format = v; end
    alias :PresentationFormat= :presentation_format=

    # Sets the lines property of your app.xml file
    def lines=(v) Axlsx::validate_int v; @lines = v; end
    alias :Lines= :lines=

    # Sets the paragraphs property of your app.xml file
    def paragraphs=(v) Axlsx::validate_int v; @paragraphs = v; end
    alias :Paragraphs= :paragraphs=

    # sets the slides property of your app.xml file
    def slides=(v) Axlsx::validate_int v; @slides = v; end
    alias :Slides= :slides=

    # sets the notes property of your app.xml file
    def notes=(v) Axlsx::validate_int v; @notes = v; end
    alias :Notes= :notes=

    # Sets the total_time property of your app.xml file
    def total_time=(v) Axlsx::validate_int v; @total_time = v; end
    alias :TotalTime= :total_time=

    # Sets the hidden_slides property of your app.xml file
    def hidden_slides=(v) Axlsx::validate_int v; @hidden_slides = v; end
    alias :HiddenSlides= :hidden_slides=

    # Sets the m_m_clips property of your app.xml file
    def m_m_clips=(v) Axlsx::validate_int v; @m_m_clips = v; end
    alias :MMClips= :m_m_clips=

    # Sets the scale_crop property of your app.xml file
    def scale_crop=(v) Axlsx::validate_boolean v; @scale_crop = v; end
    alias :ScaleCrop= :scale_crop=

    # Sets the links_up_to_date property of your app.xml file
    def links_up_to_date=(v) Axlsx::validate_boolean v; @links_up_to_date = v; end
    alias :LinksUpToDate= :links_up_to_date=

    # Sets the characters_with_spaces property of your app.xml file
    def characters_with_spaces=(v) Axlsx::validate_int v; @characters_with_spaces = v; end
    alias :CharactersWithSpaces= :characters_with_spaces=

    # Sets the share_doc property of your app.xml file
    def shared_doc=(v) Axlsx::validate_boolean v; @shared_doc = v; end
    alias :SharedDoc= :shared_doc=

    # Sets the hyperlink_base property of your app.xml file
    def hyperlink_base=(v) Axlsx::validate_string v; @hyperlink_base = v; end
    alias :HyperlinkBase= :hyperlink_base=

    # Sets the HyperLinksChanged property of your app.xml file
    def hyperlinks_changed=(v) Axlsx::validate_boolean v; @hyperlinks_changed = v; end
    alias :HyperLinksChanged= :hyperlinks_changed=

    # Sets the app_version property of your app.xml file
    def app_version=(v) Axlsx::validate_string v; @app_version = v; end
    alias :AppVersion= :app_version=

    # Sets the doc_security property of your app.xml file
    def doc_security=(v) Axlsx::validate_int v; @doc_security = v; end
    alias :DocSecurity= :doc_security=

    # Serialize the app.xml document
    # @return [String]
    def to_xml_string(str = '')
      str << '<?xml version="1.0" encoding="UTF-8"?>'
      str << '<Properties xmlns="' << APP_NS << '" xmlns:vt="' << APP_NS_VT << '">'
      instance_values.each do |key, value|
        node_name = Axlsx.camel(key)
        str << "<#{node_name}>#{value}</#{node_name}>"
      end
      str << '</Properties>'
    end

  end

end
