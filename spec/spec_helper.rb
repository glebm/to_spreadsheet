ENV["RAKE_ENV"] ||= 'test'
require 'rspec/autorun'
$: << File.expand_path('../lib', __FILE__)
require 'to_spreadsheet'
require 'haml'

RSpec.configure do |config|
  include ToSpreadsheet::Helpers
end

def build_spreadsheet(src = {})
  haml = if src[:haml]
           src[:haml]
         elsif src[:file]
           File.read(File.expand_path "support/#{src[:file]}", File.dirname(__FILE__))
         end
  ToSpreadsheet::Renderer.to_package(Haml::Engine.new(haml).render)
end