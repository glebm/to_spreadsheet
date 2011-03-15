class ContactsController < ApplicationController
  respond_to :xls, :html
  def index
    @contacts = Contact.all
    respond_with(@contacts)
  end
end