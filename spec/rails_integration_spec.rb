require 'to_spreadsheet/railtie'

describe ToSpreadsheet::Railtie do

  it "registers a renderer" do
    expect(ActionController::Renderers::RENDERERS).to include(:xlsx)
  end

end
