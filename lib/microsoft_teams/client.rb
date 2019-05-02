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

    def send_to_conversation(
      service_url:,
      conversation_id:,
      text: nil,
      attachments: nil,
      summary: nil,
      notify: false
    )
      raise ArgumentError, 'service_url should be present' if service_url.nil?
      raise ArgumentError, 'conversation_id should be present' if conversation_id.nil?

      authorized_request
        .post(
          "#{service_url}v3/conversations/#{conversation_id}/activities",
          json: {
            type: 'message',
            conversation: {
              id: conversation_id,
            },
            text: text,
            attachments: attachments,
            summary: summary,
            channelData: {
              notification: {
                alert: notify
              }
            }
          }
        ).parse
    end

    def reply_to_activity(
      service_url:,
      conversation_id:,
      activity_id:,
      text: nil,
      attachments: nil,
      summary: nil,
      notify: false
    )
      raise ArgumentError, 'service_url should be present' if service_url.nil?
      raise ArgumentError, 'conversation_id should be present' if conversation_id.nil?
      raise ArgumentError, 'activity_id should be present' if activity_id.nil?

      authorized_request
        .post(
          "#{service_url}v3/conversations/#{conversation_id}/activities/#{activity_id}",
          json: {
            type: 'message',
            conversation: {
              id: conversation_id,
            },
            replyToId: activity_id,
            text: text,
            attachments: attachments,
            summary: summary,
            channelData: {
              notification: {
                alert: notify
              }
            }
          }
        ).parse
    end

    def get_direct_conversation_id(service_url:, tenant_id:, bot_id:, user_id:)
      raise ArgumentError, 'service_url should be present' if service_url.nil?
      raise ArgumentError, 'tenant_id should be present' if tenant_id.nil?
      raise ArgumentError, 'bot_id should be present' if bot_id.nil?
      raise ArgumentError, 'user_id should be present' if user_id.nil?

      response = authorized_request
        .post(
          "#{service_url}v3/conversations",
          json: {
            bot: {
              id: bot_id,
              name: Flek.env.microsoft_teams_bot_name
            },
            members: [
              {
                id: user_id
              }
            ],
            channelData: {
              tenant: {
                id: tenant_id
              }
            }
          }
        ).parse

      response['id']
    end

    def send_to_connector(webhook_url:, message:)
      raise ArgumentError, 'webhook_url should be present' if webhook_url.nil?
      raise ArgumentError, 'message should be present' if message.nil?

      # This API returns 1 is successful and the error message otherwise
      request
        .post(
          webhook_url,
          json: message
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

      if response['error'].present?
        raise AuthenticationError,
          "#{response['error_codes'].inspect} - #{response['error']}: #{response['error_description']}"
      end

      response['access_token']
    end
  end
end
