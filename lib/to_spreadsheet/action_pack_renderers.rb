require 'active_support'
require 'action_controller/metal/renderers'
require 'action_controller/metal/responder'

require 'to_spreadsheet/xlsx'

# This will let us do thing like `render :xlsx => 'index'`
# This is also how Rails internally implements its :json and :xml renderers
# Rarely used, nevertheless public API
ActionController::Renderers.add :xlsx do |template, options|
  filename = options[:filename] || options[:template] || 'data'
  send_data ToSpreadsheet::XLSX.to_io(render_to_string(options[:template], options)).read, :type => :xlsx, :disposition => "attachment; filename=\"#{filename}.xlsx\""
end

class ActionController::Responder
  # This sets up a default render call for when you do
  # respond_to do |format|
  #   format.xls
  # end
  def to_xlsx
    controller.render :xlsx => controller.action_name
  end
end
