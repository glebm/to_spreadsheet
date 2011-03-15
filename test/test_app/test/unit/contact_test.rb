require 'test_helper'

class ContactTest < ActiveSupport::TestCase
  def test_contact
    assert_equal 1, Contact.count
  end
end