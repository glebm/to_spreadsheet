# encoding: UTF-8
require 'rubygems'
require 'bundler/gem_tasks'

require 'rake'
require 'rdoc/task'
require 'rspec/core/rake_task'
RSpec::Core::RakeTask.new(:spec)

task :env do
  $: << File.expand_path('lib', File.dirname(__FILE__))
  require 'to_spreadsheet'
  include ToSpreadsheet::Helpers
end

desc 'Generate a simple xlsx file'
task :write_test_xlsx => :env do
  require 'haml'
  path = '/tmp/spreadsheet.xlsx'
  html = Haml::Engine.new(File.read('spec/support/table.html.haml')).render
  ToSpreadsheet::Renderer.to_package(html).serialize(path)
  puts "Written to #{path}"
end

desc "Test all rails versions"
task :test_all_rails do
  Dir.glob("./spec/gemfiles/*{[!.lock]}").each do |gemfile|
    puts "TESTING WITH #{gemfile}"
    system "BUNDLE_GEMFILE=#{gemfile} bundle | grep Installing"
    system "BUNDLE_GEMFILE=#{gemfile} bundle exec rspec"
  end
end