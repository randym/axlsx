    require 'rubygems'
    require 'nokogiri'
    require 'open-uri'
    require 'axlsx'

    class Scraper

       def initialize(url, selector)
         @url = url
         @selector = selector
       end

       def hooks
         @hooks ||= {}
       end

       def add_hook(clue, p_roc)
         hooks[clue] = p_roc
       end

       def export(file_name)
         Scraper.clues.each do |clue|
           if detail = parse_clue(clue)
             output << [clue, detail.pop]
             detail.each { |datum| output << ['', datum] }
           end
         end
         serialize(file_name)
       end

       private

       def self.clues
         @clues ||= ['Operating system', 'Processors', 'Chipset', 'Memory type', 'Hard drive', 'Graphics',
                     'Ports', 'Webcam', 'Pointing device', 'Keyboard', 'Network interface', 'Chipset', 'Wireless',
                     'Power supply type', 'Energy efficiency', 'Weight', 'Minimum dimensions (W x D x H)',
                     'Warranty', 'Software included', 'Product color']
       end

       def doc
         @doc ||= begin 
                    Nokogiri::HTML(open(@url))
                  rescue
                    raise ArgumentError, 'Invalid URL - Nothing to parse'
                  end
       end

       def output
         @output ||= []
       end

       def selector_for_clue(clue)
         @selector % clue
       end

       def parse_clue(clue)
         if element = doc.at(selector_for_clue(clue))
           call_hook(clue, element) || element.inner_html.split('<br>').each(&:strip)
         end
       end

       def call_hook(clue, element)
         if hooks[clue].is_a? Proc
            value = hooks[clue].call(element)
            value.is_a?(Array) ? value : [value]
         end
       end

       def package
         @package ||= Axlsx::Package.new
       end

       def serialize(file_name)
         package.workbook.add_worksheet do |sheet|
           output.each { |datum| sheet.add_row datum }
         end
         package.serialize(file_name)
       end
    end

    scraper = Scraper.new("http://h10010.www1.hp.com/wwpc/ie/en/ho/WF06b/321957-321957-3329742-89318-89318-5186820-5231694.html?dnr=1", "//td[text()='%s']/following-sibling::td")

    # define a custom action to take against any elements found.
    os_parse = Proc.new do |element|
      element.inner_html.split('<br>').each(&:strip!).each(&:upcase!)
    end

    scraper.add_hook('Operating system', os_parse)

    scraper.export('foo.xlsx')

