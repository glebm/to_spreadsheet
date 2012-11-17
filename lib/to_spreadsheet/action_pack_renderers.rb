require 'active_support'
require 'action_controller/metal/renderers'
require 'action_controller/metal/responder'

require 'to_spreadsheet/renderer'

# This will let us do thing like `render :xlsx => 'index'`
# This is similar to how Rails internally implements its :json and :xml renderers
ActionController::Renderers.add :xlsx do |template, options|
  filename = options[:filename] || options[:template] || 'data'

  html = with_context ToSpreadsheet::Context.global.merge(ToSpreadsheet::Context.new) do
    # local context
    @local_formats.each do |selector, &block|
      context.process_dsl selector, &block
    end if @local_formats
    render_to_string(options[:template], options)
  end

  data = ToSpreadsheet::Axlsx::Renderer.to_data(html)
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
