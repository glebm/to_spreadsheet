require 'active_support'
require 'action_controller/metal/renderers'
require 'action_controller/metal/responder'

require 'to_spreadsheet/xls'

# This will let us do thing like `render :xls => 'index'`
# This is also how Rails internally implements its :json and :xml renderers
# Rarely used, nevertheless public API
ActionController::Renderers.add :xls do |template, options|
  send_data ToSpreadsheet::XLS.to_io(render_to_string(template, options)).read, :type => :xls
end

class ActionController::Responder
  # This sets up a default render call for when you do
  # respond_to do |format|
  #   format.xls
  # end
  def to_xls
    controller.render :xls => controller.action_name
  end
end