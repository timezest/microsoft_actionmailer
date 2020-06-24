require "microsoft_actionmailer/version"
require 'microsoft_actionmailer/railtie' if defined?(Rails)
require 'microsoft_actionmailer/api'

require 'httparty'
require 'net/http'
require 'uri'

module MicrosoftActionmailer

  GRAPH_HOST = 'https://graph.microsoft.com'.freeze

  class ApiError < StandardError; end

  class DeliveryMethod
    include MicrosoftActionmailer::Api

    attr_reader :access_token
    attr_reader :delivery_options

    def initialize params
      @access_token = params[:authorization]
      @delivery_options = params[:delivery_options] || {}
    end

    def deliver! mail
      body = if mail.html_part.present?
               mail.html_part.body.encoded
             else
               mail.body.encoded
             end

      before_send = delivery_options[:before_send]
      if before_send && before_send.respond_to?(:call)
        before_send.call(mail)
      end

      message = ms_send_mail(
        access_token,
        mail.subject,
        body,
        mail.from,
        mail.to,
        mail.attachments,
        mail.reply_to
      )

      after_send = delivery_options[:after_send]
      if after_send && after_send.respond_to?(:call)
        after_send.call(mail)
      end
    end
  end
end
