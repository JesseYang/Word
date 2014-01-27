require 'rubygems'
require 'sinatra'
require File.expand_path '../parse_word.rb', __FILE__
#\ -w -p 9292
# map 'http://export.diaoyan.me:9292/' do
run Sinatra::Application
# end
