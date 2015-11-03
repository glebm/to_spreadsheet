require 'rails/railtie'
module ToSpreadsheet
  class Railtie < ::Rails::Railtie
    config.to_spreadsheet = ActiveSupport::OrderedOptions.new
    config.to_spreadsheet.renderer = ToSpreadsheet.renderer

    config.after_initialize do |app|
      ToSpreadsheet.instance_variable_set("@renderer", app.config.to_spreadsheet.renderer)

      require 'to_spreadsheet/rails/action_pack_renderers'
      require 'to_spreadsheet/rails/view_helpers'
      require 'to_spreadsheet/rails/mime_types'
      ActionView::Base.send :include, ToSpreadsheet::Rails::ViewHelpers
    end
  end
end