module MicrosoftActionmailer
  module Api
    # Creates a message and saves in 'Draft' mailFolder
    def ms_create_message(token, subject, content, address)
      query = {
        "subject": subject,
        "importance": "Normal",
        "body":{
          "contentType": "HTML",
          "content": content
        },
        "toRecipients": [
          {
            "emailAddress": {
              "address": address
            }
          }
        ]
      }
      create_message_url = '/v1.0/me/messages'
      req_method = 'post'
      response = make_api_call create_message_url, token, query,req_method
      raise response.parsed_response.to_s || "Request returned #{response.code}" unless response.code == 201
      response
    end

    # Sends the message created using message id
    def ms_send_message(token, message_id)
      send_message_url = "/v1.0/me/messages/#{message_id}/send"
      req_method = 'post'
      query = {}
      response = make_api_call send_message_url, token, query,req_method
      raise response.parsed_response.to_s || "Request returned #{response.code}" unless response.code == 202
      response
    end

    def make_api_call(token)
      Excon.post('https://graph.microsoft.com/v1.0/me/sendMail',
                 body: {
                   'message': {
                     'subject': 'Meet for lunch?',
                     'body': {
                       'contentType': 'Text',
                       'content': 'The new cafeteria is open.'
                     },
                     'from': {
                       'emailAddress': {
                         'address': 'inbox@codeultras.com'
                       }
                     },
                     'toRecipients': [
                       {
                         'emailAddress': {
                           'address': 'artem.kulakov@gmail.com'
                         }
                       }
                     ]
                   },
                   'saveToSentItems': 'false'
                 }.to_json,
                 headers: {
                   'Content-Type' => 'application/json',
                   'Authorization' => "Bearer #{token}"
                 }
      )
    end
  end
end
