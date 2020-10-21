module MicrosoftActionmailer
  module Api
    # Sends the mail directly
    def ms_send_mail(token, proxy, subject, content, from_address, addresses, attachments, reply_to)
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

      reply_to_recipient = reply_to.present? ? [{ "emailAddress": { "address": reply_to.first } }] : []

      query = { "message": {
        "subject": subject,
        "importance": "Normal",
        "body":{
          "contentType": "HTML",
          "content": content
        },
        "toRecipients": recipients,
        "replyTo": reply_to_recipient,
        "from": { "emailAddress": { "address": from_address.first } },
        "attachments": attachment_list
      }}


      response = make_api_call('/v1.0/me/sendMail', token, proxy, query, :post)

      raise ApiError.new(JSON.parse(response.body)) || "Request returned #{response.code}" unless response.status == 202
      response
    end

    def make_api_call(endpoint, token, proxy, params = nil, req_method)
      connection = Excon.new(GRAPH_HOST, proxy.present? ? { proxy: proxy, debug: true } : { debug: true})
      headers = {
        'Authorization'=> "Bearer #{token}",
        'Content-Type' => 'application/json'
      }

      query = params || {}
      if req_method == :get
        connection.get path: endpoint,
                   headers: headers,
                   query: query
      elsif req_method == :post
        connection.post path: endpoint,
                   headers: headers,
                   body: query.to_json
      end
    end
  end
end
