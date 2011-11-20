# -*- coding: utf-8 -*-
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

    # @return [String] The name of the document template.
    attr_accessor :Template

    # @return [String] The name of the manager for the document.
    attr_accessor :Manager

    # @return [String] The name of the company generating the document.
    attr_accessor :Company 

    # @return [Integer] The number of pages in the document.
    attr_accessor :Pages

    # @return [Integer] The number of words in the document.
    attr_accessor :Words

    # @return [Integer] The number of characters in the document.
    attr_accessor :Characters

    # @return [String] The intended format of the presentation.
    attr_accessor :PresentationFormat

    # @return [Integer] The number of lines in the document.
    attr_accessor :Lines

    # @return [Integer] The number of paragraphs in the document
    attr_accessor :Paragraphs

    # @return [Intger] The number of slides in the document.
    attr_accessor :Slides

    # @return [Integer] The number of slides that have notes.
    attr_accessor :Notes

    # @return [Integer] The total amount of time spent editing.   
    attr_accessor :TotalTime

    # @return [Integer] The number of hidden slides.
    attr_accessor :HiddenSlides

    # @return [Integer] The total number multimedia clips
    attr_accessor :MMClips

    # @return [Boolean] The display mode for the document thumbnail.
    attr_accessor :ScaleCrop

    # @return [Boolean] The links in the document are up to date.
    attr_accessor :LinksUpToDate

    # @return [Integer] The number of characters in the document including spaces.
    attr_accessor :CharactersWithSpaces
    
    # @return [Boolean] Indicates if the document is shared.
    attr_accessor :ShareDoc

    # @return [String] The base for hyper links in the document. 
    attr_accessor :HyperLinkBase

    # @return [Boolean] Indicates that the hyper links in the document have been changed.
    attr_accessor :HyperlinksChanged

    # @return [String] The name of the application
    attr_accessor :Application

    # @return [String] The version of the application.
    attr_accessor :AppVersion

    # @return [Integer] Document security
    attr_accessor :DocSecurity

    # Creates an App object
    # @option options [String] Template
    # @option options [String] Manager
    # @option options [Integer] Pages
    # @option options [Integer] Words
    # @option options [Integer] Characters
    # @option options [String] PresentationFormat
    # @option options [Integer] Lines
    # @option options [Integer] Paragraphs
    # @option options [Integer] Slides
    # @option options [Integer] Notes
    # @option options [Integer] TotalTime
    # @option options [Integer] HiddenSlides
    # @option options [Integer] MMClips
    # @option options [Boolean] ScaleCrop
    # @option options [Boolean] LinksUpToDate
    # @option options [Integer] CharactersWithSpaces
    # @option options [Boolean] ShareDoc    
    # @option options [String] HyperLinkBase
    # @option options [String] HyperlinksChanged
    # @option options [String] Application
    # @option options [String] AppVersion
    # @option options [Integer] DocSecurity
    def initalize(options={})
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? o[0]
      end
    end

    def Template=(v) Axlsx::validate_string v; @Template = v; end 
    def Manager=(v) Axlsx::validate_string v; @Manager = v; end 
    def Company=(v) Axlsx::validate_string v; @Company = v; end 
    def Pages=(v) Axlsx::validate_int v; @Pages = v; end 
    def Words=(v) Axlsx::validate_int v; @Words = v; end 
    def Characters=(v) Axlsx::validate_int v; @Characters = v; end 
    def PresentationFormat=(v) Axlsx::validate_string v; @PresentationFormat = v; end 
    def Lines=(v) Axlsx::validate_int v; @Lines = v; end 
    def Paragraphs=(v) Axlsx::validate_int v; @Paragraphs = v; end 
    def Slides=(v) Axlsx::validate_int v; @Slides = v; end 
    def Notes=(v) Axlsx::validate_int v; @Notes = v; end 
    def TotalTime=(v) Axlsx::validate_int v; @TotalTime = v; end 
    def HiddenSlides=(v) Axlsx::validate_int v; @HiddenSlides = v; end 
    def MMClips=(v) Axlsx::validate_int v; @MMClips = v; end 
    def ScaleCrop=(v) Axlsx::validate_boolean v; @ScaleCrop = v; end 
    def LinksUpToDate=(v) Axlsx::validate_boolean v; @LinksUpToDate = v; end 
    def CharactersWithSpaces=(v) Axlsx::validate_int v; @CharactersWithSpaces = v; end 
    def ShareDoc=(v) Axlsx::validate_boolean v; @ShareDoc = v; end 
    def HyperLinkBase=(v) Axlsx::validate_string v; @HyperLinkBase = v; end 
    def HyperlinksChanged=(v) Axlsx::validate_boolean v; @HyperlinksChanged = v; end 
    def Application=(v) Axlsx::validate_string v; @Application = v; end 
    def AppVersion=(v) Axlsx::validate_string v; @AppVersion = v; end 
    def DocSecurity=(v) Axlsx::validate_int v; @DocSecurity = v; end 

    # Generate an app.xml document
    # @return [String] The document as a string
    def to_xml()
      builder = Nokogiri::XML::Builder.new(:encoding => ENCODING) do |xml|
        xml.send(:Properties, :xmlns => APP_NS, :'xmlns:vt' => APP_NS_VT) {
          self.instance_values.each do |name, value|
            xml.send("ap:#{name}", value)
          end
        }
      end      
      builder.to_xml
    end
  end
end
