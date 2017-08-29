module.exports = {
    creds: {
      redirectUrl: 'http://localhost:3000/login',
      clientID: '<clientID>',
      clientSecret: '<clientSecret>',
      identityMetadata: 'https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration',
      allowHttpForRedirectUrl: true, // For development only
      responseType: 'code',
      validateIssuer: false, // For development only
      responseMode: 'query',
      scope: ['User.Read', 'Mail.Send', 'Files.ReadWrite', 'Calendars.ReadWrite', 'Directory.AccessAsUser.All']
    }
  };
  