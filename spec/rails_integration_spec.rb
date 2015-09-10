require "to_spreadsheet/railtie"

describe ToSpreadsheet::Railtie do

  it "should register renderer" do
    expect(ActionController::Renderers::RENDERERS).to include(:xlsx)
  end

end
