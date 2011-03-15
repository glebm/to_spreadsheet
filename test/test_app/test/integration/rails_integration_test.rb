require 'test_helper'

class RailsIntegrationTest < ActionController::IntegrationTest
  fixtures :all
  test "html" do
    get '/contacts'
    assert_response :success
  end

  test "xls" do
    get '/contacts.xls'
    assert_response :success
  end
end