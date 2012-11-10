# encoding: UTF-8
require 'rubygems'
require 'bundler/setup'

require 'rake'
require 'rdoc/task'
require 'rspec/core/rake_task'
RSpec::Core::RakeTask.new(:spec)

task :env do
  $: << File.expand_path('lib', File.dirname(__FILE__))
  require 'to_spreadsheet'
  include ToSpreadsheet::Helpers
end

task :write_test_xlsx => :env do
  require 'haml'
  path = '/tmp/spreadsheet.xlsx'
  html = Haml::Engine.new(File.read('spec/support/table.html.haml')).render
  ToSpreadsheet::Axlsx::Renderer.to_package(html).serialize(path)
  puts "Written to #{path}"
end