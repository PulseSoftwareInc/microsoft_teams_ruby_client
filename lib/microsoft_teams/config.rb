module MicrosoftTeams
  module Config
    extend self

    ATTRIBUTES = %i[
      access_token
      app_id
      app_password
    ].freeze

    attr_accessor(*Config::ATTRIBUTES)

    def reset
      Config::ATTRIBUTES.each do |key|
        send("#{key}=", nil)
      end
    end

    reset
  end

  class << self
    def configure
      block_given? ? yield(Config) : Config
    end

    def config
      Config
    end
  end
end
