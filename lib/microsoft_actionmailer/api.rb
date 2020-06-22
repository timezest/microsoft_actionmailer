module MicrosoftActionmailer
  module Api
    # Sends the mail directly
    def ms_send_mail(token, subject, content, from_address, addresses, attachments)
      attachment_list = []
      attachments.each do |attachment|
        data = { "@odata.type": "#microsoft.graph.fileAttachment",
                 "name": attachment.filename,
                 "contentType": attachment.content_type,
                 "contentBytes": Base64.encode64(attachment.body.raw_source)
               }
        attachment_list << data
      end

      recipients = []
      addresses.each do |address|
        data = { "emailAddress": { "address": address } }
        recipients << data
      end

      query = { "message": {
        "subject": subject,
        "importance": "Normal",
        "body":{
          "contentType": "HTML",
          "content": content
        },
        "toRecipients": recipients,
        "from": { "emailAddress": { "address": from_address.first } },
        "attachments": attachment_list
      }}

      response = make_api_call('/v1.0/me/sendMail', token, query, :post)

      raise ApiError.new(response.parsed_response.to_s) || "Request returned #{response.code}" unless response.code == 202
      response
    end

    # Creates a message and saves in 'Draft' mailFolder
    def ms_create_message(token, subject, content, addresses, attachments)

      attachment_list = []
      attachments.each do |attachment|
        data = { "@odata.type": "#microsoft.graph.fileAttachment",
                 "name": attachment.filename,
                 "contentType": attachment.content_type,
                 "contentBytes": Base64.encode64(attachment.body.raw_source)
               }
        attachment_list << data
      end

      recipients = []
      addresses.each do |address|
        data = { "emailAddress": { "address": address } }
        recipients << data
      end

      query = {
        "subject": subject,
        "importance": "Normal",
        "body":{
          "contentType": "HTML",
          "content": content
        },
        "toRecipients": recipients,
        "attachments": attachment_list
      }

      response = make_api_call('/v1.0/me/messages', token, query, :post)
      raise response.parsed_response.to_s || "Request returned #{response.code}" unless response.code == 201
      response
    end

    # Sends the message created using message id
    def ms_send_message(token, message_id)
      send_message_url = "/v1.0/me/messages/#{message_id}/send"
      response = make_api_call(send_message_url, token, {}, :post)
      raise response.parsed_response.to_s || "Request returned #{response.code}" unless response.code == 202
      response
    end

    def make_api_call(endpoint, token, params = nil, req_method)
      headers = {
        'Authorization'=> "Bearer #{token}",
        'Content-Type' => 'application/json'
      }

      query = params || {}
      if req_method == :get
        HTTParty.get "#{GRAPH_HOST}#{endpoint}",
                   headers: headers,
                   query: query
      elsif req_method == :post
        HTTParty.post "#{GRAPH_HOST}#{endpoint}",
                   headers: headers,
                   body: query.to_json
      end
    end
  end
end
