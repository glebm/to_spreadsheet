require 'to_spreadsheet/railtie'

describe ToSpreadsheet::Railtie do

  it "registers a renderer" do
    expect(ToSpreadsheet.renderer).to eq(:html2xlsx)
    expect(ActionController::Renderers::RENDERERS).to include(ToSpreadsheet.renderer)
  end

end
