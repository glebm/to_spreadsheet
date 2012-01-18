ENV["RAKE_ENV"] ||= 'test'
require 'rspec/autorun'
$: << File.expand_path('../lib', __FILE__)
require 'to_spreadsheet'
require 'haml'