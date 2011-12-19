class CreateContacts < ActiveRecord::Migration
  class << self
    def up
      create_table :contacts do |t|
        t.string :name
        t.string :website
        t.integer :age
      end
    end

    def down
      drop_table :contacts
    end
  end
end