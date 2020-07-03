const { API_BASE_URL } = require('./constants');
const AdmZip = require('adm-zip');

const zipFiles = files => {
  let zip = new AdmZip();
  let size = 0;

  files.forEach(file => {
    zip.addFile(file.name, new Buffer(file.contentBytes, 'base64'));
    size += file.size;
  });

  return {
    zip,
    size,
  };
};

async function getUserPrincipalName(z, bundle) {
  const getUserPrincipalNameError =
    'There was an error obtaining necessary user information. Please try reconnecting your account or contact Zapier support for assistance.';

  if (bundle.authData.userPrincipalName) {
    return bundle.authData.userPrincipalName;
  }

  const response = await z.request({
    method: 'GET',
    url: `${API_BASE_URL}/me`,
    prefixErrorMessageWith: getUserPrincipalNameError,
  });

  if (!response || !response.content) {
    z.console.log(
      'Unable to obtain user principal name. Did not receive a valid response from Microsoft.'
    );
    throw new Error(getUserPrincipalNameError);
  }

  const parsedJSON = z.JSON.parse(response.content);
  if (parsedJSON && parsedJSON.userPrincipalName) {
    bundle.authData.userPrincipalName = parsedJSON.userPrincipalName; // eslint-disable-line no-param-reassign
    return parsedJSON.userPrincipalName;
  }

  z.console.log(
    `Unable to obtain user principal name. Parsed response is: ${parsedJSON}`
  );
  throw new Error(getUserPrincipalNameError);
}

module.exports = {
  getUserPrincipalName,
};
