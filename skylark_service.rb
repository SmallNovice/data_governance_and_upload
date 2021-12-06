require 'jwt'
require 'rest-client'

class SkylarkService
  attr_reader :appid, :appsecret, :namespace_id, :host

  def initialize
    @namespace_id = 1
    @appid = 'f4f34a327fb4e2d5e87e5622f8ebb4cc45c7a8212650ebf2b95314f5170f0418'
    @appsecret = '7b99f2c7f85f148c57d9ed4cd1bdcd3d59a99a1666b0af6395f8fd1d22f2f001'
    @host = 'https://gxhz.cdht.gov.cn'
  end

  def get_form(form_id)
    RestClient::Request.execute(
      method: :get,
      url: get_form_url(form_id),
      headers: authorization_token,
      timeout: 3
    )
  end

  private

  def authorization_token
    { Authorization: "#{@appid}:#{encode_secret}" }
  end

  def encode_secret
    JWT.encode(
      {
        namespace_id: @namespace_id,
      },
      @appsecret,
      'HS256',
      typ: 'JWT', alg: 'HS256'
    )
  end

  def get_form_url(form_id)
    "#{@host}/api/v4/forms/#{form_id}"
  end
end

