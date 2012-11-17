require 'action_dispatch/http/mime_type'
Mime::Type.register "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", :xlsx unless Mime::Type.lookup_by_extension(:xlsx)
