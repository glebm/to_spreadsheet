require 'rails/railtie'
require 'to_spreadsheet/rails/mime_types'
require 'to_spreadsheet/rails/action_pack_renderers'
require 'to_spreadsheet/rails/view_helpers'
module ToSpreadsheet
  class Railtie < ::Rails::Railtie
    initializer "to_spreadsheet.view_helpers" do
      ActionView::Base.send :include, ToSpreadsheet::Rails::ViewHelpers
    end
  end
end