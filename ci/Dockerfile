# vim: set ft=Dockerfile

FROM vindvaki/xlsx-validator

RUN apt-get install -y build-essential ruby ruby-dev libxslt-dev libxml2-dev zlib1g-dev
RUN gem install bundler
ADD Gemfile* axlsx.gemspec /app/axlsx/
WORKDIR /app/axlsx/
ADD lib/axlsx/version.rb lib/axlsx/version.rb
RUN bundle install
ADD . .

ADD ci/validate-xlsx-files.sh .
