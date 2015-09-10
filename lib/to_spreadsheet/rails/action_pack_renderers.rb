require 'active_support'
require 'action_controller/metal/renderers'

# in rails 3.2  it's ActiveSupport::VERSION
# in rails 4.0+ it's ActiveSupport.version (instance of Gem::Version)
if ActiveSupport.respond_to?(:version) && ActiveSupport.version.to_s >= "4.2.0"
  # For rails 4.2
  require 'action_controller/responder'
else
  # For rails 3.2 - rails 4.1
  require 'action_controller/metal/responder'
end


# This will let us do thing like `render :xlsx => 'index'`
# This is similar to how Rails internally implements its :json and :xml renderers
ActionController::Renderers.add :xlsx do |template, options|
  filename = options[:filename] || options[:template] || 'data'
  data = ToSpreadsheet::Context.with_context ToSpreadsheet::Context.global.merge(ToSpreadsheet::Context.new) do |context|
    html = render_to_string(template, options.merge(template: template.to_s, formats: ['xlsx', 'html']))
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
