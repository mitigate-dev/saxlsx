require 'bigdecimal'
require 'rational'
require 'zip'
require 'ox'
require 'cgi'

Dir["#{File.dirname(__FILE__)}/saxlsx/**/*.rb"].each { |f| require f }
