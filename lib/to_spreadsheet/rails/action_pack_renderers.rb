require 'active_support'
require 'action_controller/metal/renderers'
require 'action_controller/metal/responder'

# This will let us do thing like `render :xlsx => 'index'`
# This is similar to how Rails internally implements its :json and :xml renderers
ActionController::Renderers.add :xlsx do |template, options|
  filename = options[:filename] || options[:template] || 'data'
  data = ToSpreadsheet::Context.with_context ToSpreadsheet::Context.global.merge(ToSpreadsheet::Context.new) do |context|
    html = render_to_string(template, options.merge(template: template.to_s, formats: ['html']))
    ToSpreadsheet::Renderer.to_data(html, context)
  end
  send_data data, type: :xlsx, disposition: %(attachment; filename="#{filename}.xlsx")
end

class ActionController::Responder
  # This sets up a default render call for when you do
  # respond_to do |format|
  #   format.xlsx
  # end
  def to_xlsx
    controller.render xlsx: controller.action_name
  end
end
