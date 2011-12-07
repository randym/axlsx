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
    attr_reader :Template

    # @return [String] The name of the manager for the document.
    attr_reader :Manager

    # @return [String] The name of the company generating the document.
    attr_reader :Company 

    # @return [Integer] The number of pages in the document.
    attr_reader :Pages

    # @return [Integer] The number of words in the document.
    attr_reader :Words

    # @return [Integer] The number of characters in the document.
    attr_reader :Characters

    # @return [String] The intended format of the presentation.
    attr_reader :PresentationFormat

    # @return [Integer] The number of lines in the document.
    attr_reader :Lines

    # @return [Integer] The number of paragraphs in the document
    attr_reader :Paragraphs

    # @return [Intger] The number of slides in the document.
    attr_reader :Slides

    # @return [Integer] The number of slides that have notes.
    attr_reader :Notes

    # @return [Integer] The total amount of time spent editing.   
    attr_reader :TotalTime

    # @return [Integer] The number of hidden slides.
    attr_reader :HiddenSlides

    # @return [Integer] The total number multimedia clips
    attr_reader :MMClips

    # @return [Boolean] The display mode for the document thumbnail.
    attr_reader :ScaleCrop

    # @return [Boolean] The links in the document are up to date.
    attr_reader :LinksUpToDate

    # @return [Integer] The number of characters in the document including spaces.
    attr_reader :CharactersWithSpaces
    
    # @return [Boolean] Indicates if the document is shared.
    attr_reader :ShareDoc

    # @return [String] The base for hyper links in the document. 
    attr_reader :HyperLinkBase

    # @return [Boolean] Indicates that the hyper links in the document have been changed.
    attr_reader :HyperlinksChanged

    # @return [String] The name of the application
    attr_reader :Application

    # @return [String] The version of the application.
    attr_reader :AppVersion

    # @return [Integer] Document security
    attr_reader :DocSecurity

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
    def initialize(options={})
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
    end

    # Sets the Template property of your app.xml file
    def Template=(v) Axlsx::validate_string v; @Template = v; end 

    # Sets the Manager property of your app.xml file
    def Manager=(v) Axlsx::validate_string v; @Manager = v; end 

    # Sets the Company property of your app.xml file
    def Company=(v) Axlsx::validate_string v; @Company = v; end 

    # Sets the Pages property of your app.xml file
    def Pages=(v) Axlsx::validate_int v; @Pages = v; end 

    # Sets the Words property of your app.xml file
    def Words=(v) Axlsx::validate_int v; @Words = v; end 

    # Sets the Characters property of your app.xml file
    def Characters=(v) Axlsx::validate_int v; @Characters = v; end 


    # Sets the PresentationFormat property of your app.xml file
    def PresentationFormat=(v) Axlsx::validate_string v; @PresentationFormat = v; end 
    # Sets the Lines property of your app.xml file
    def Lines=(v) Axlsx::validate_int v; @Lines = v; end 
    # Sets the Paragraphs property of your app.xml file
    def Paragraphs=(v) Axlsx::validate_int v; @Paragraphs = v; end 
    # Sets the Slides property of your app.xml file
    def Slides=(v) Axlsx::validate_int v; @Slides = v; end 
    # Sets the Notes property of your app.xml file
    def Notes=(v) Axlsx::validate_int v; @Notes = v; end 
    # Sets the TotalTime property of your app.xml file
    def TotalTime=(v) Axlsx::validate_int v; @TotalTime = v; end 
    # Sets the HiddenSlides property of your app.xml file
    def HiddenSlides=(v) Axlsx::validate_int v; @HiddenSlides = v; end 
    # Sets the MMClips property of your app.xml file
    def MMClips=(v) Axlsx::validate_int v; @MMClips = v; end 
    # Sets the ScaleCrop property of your app.xml file
    def ScaleCrop=(v) Axlsx::validate_boolean v; @ScaleCrop = v; end 
    # Sets the LinksUpToDate property of your app.xml file
    def LinksUpToDate=(v) Axlsx::validate_boolean v; @LinksUpToDate = v; end 
    # Sets the CharactersWithSpaces property of your app.xml file
    def CharactersWithSpaces=(v) Axlsx::validate_int v; @CharactersWithSpaces = v; end 
    # Sets the ShareDoc property of your app.xml file
    def ShareDoc=(v) Axlsx::validate_boolean v; @ShareDoc = v; end 
    # Sets the HyperLinkBase property of your app.xml file
    def HyperLinkBase=(v) Axlsx::validate_string v; @HyperLinkBase = v; end 
    # Sets the HyperLinksChanged property of your app.xml file
    def HyperlinksChanged=(v) Axlsx::validate_boolean v; @HyperlinksChanged = v; end 
    # Sets the Application property of your app.xml file
    def Application=(v) Axlsx::validate_string v; @Application = v; end 
    # Sets the AppVersion property of your app.xml file
    def AppVersion=(v) Axlsx::validate_string v; @AppVersion = v; end 
    # Sets the DocSecurity property of your app.xml file
    def DocSecurity=(v) Axlsx::validate_int v; @DocSecurity = v; end 

    # Generate an app.xml document
    # @return [String] The document as a string
    def to_xml()
      builder = Nokogiri::XML::Builder.new(:encoding => ENCODING) do |xml|
        xml.send(:Properties, :xmlns => APP_NS, :'xmlns:vt' => APP_NS_VT) {
          self.instance_values.each do |name, value|
            xml.send(name, value)
          end
        }
      end      
      builder.to_xml
    end
  end
end
