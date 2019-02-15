module MicrosoftTeams
  class Client
    attr_accessor(*Config::ATTRIBUTES)

    def initialize(options = {})
      Config::ATTRIBUTES.each do |key|
        send("#{key}=", options[key] || MicrosoftTeams.config.send(key))
      end

      if access_token
        @access_token = access_token
      elsif (app_id && app_password)
        @access_token = get_access_token(app_id, app_password)
      else
        raise ConfigurationError
      end
    end

    def get_team_members(service_url:, team_id:)
      authorized_request
        .get(
          "#{service_url}v3/conversations/#{team_id}/members"
        ).parse
    end

    def send_message(service_url:, conversation_id:, activity_id:, from:, recipient:, conversation:, text: nil, attachments: nil)
      authorized_request
        .post(
          "#{service_url}v3/conversations/#{conversation_id}/activities/#{activity_id}",
          json: {
            type: 'message',
            from: from,
            recipient: recipient,
            conversation: conversation,
            replyToId: activity_id,
            text: text,
            attachments: attachments,
          }
        ).parse
    end

    def send_message_to_channel(webhook_url:, text:)
      # This API returns 1 is successful and the error message otherwise
      request
        .post(
          webhook_url,
          json: {
            text: text
          }
        )
    end

    private

    def request
      HTTP.accept('application/json')
    end

    def authorized_request
      # TODO: Add handling in case access token somehow gets invalidated
      request.auth("Bearer #{@access_token}")
    end

    def get_access_token(client_id, client_secret)
      # TODO: Add handling in case this request fails
      response = HTTP.headers(accept: 'application/x-www-form-urlencoded')
        .post(
          'https://login.microsoftonline.com/botframework.com/oauth2/v2.0/token',
          form: {
            grant_type: 'client_credentials',
            scope: 'https://api.botframework.com/.default',
            client_id: client_id,
            client_secret: client_secret,
          }
        ).parse

      response['access_token']
    end
  end
end
