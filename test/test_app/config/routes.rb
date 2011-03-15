TestApp::Application.routes.draw do
  resources :contacts
  root :to => redirect("/contacts.html")
end
