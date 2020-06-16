require "microsoft_actionmailer/version"
require 'microsoft_actionmailer/railtie' if defined?(Rails)
require 'microsoft_actionmailer/api'

require 'httparty'
require 'net/http'
require 'uri'

module MicrosoftActionmailer

  GRAPH_HOST = 'https://graph.microsoft.com'.freeze

  class DeliveryMethod
    include MicrosoftActionmailer::Api

    attr_reader :access_token
    attr_reader :delivery_options

    def initialize params
      @access_token = params[:authorization]
      @delivery_options = params[:delivery_options] || {}
    end

    def deliver! mail
      res = make_api_call(access_token)
    end
  end
end
