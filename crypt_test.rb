require 'axlsx.rb'
f = 'crypt_test.xlsx'
p = Axlsx::Package.new
p.serialize f
p.encrypt f, 'password'
