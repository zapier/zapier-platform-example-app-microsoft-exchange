/**
 * Middleware allows you to run some kind of work before and after HTTP requests.
 * A beforeRequest middleware function takes a request options object, and returns a (possibly mutated) request object.
 * An afterResponse middleware function takes a response object, and returns a (possibly mutated) response object.
 *
 * Relevant docs: https://platform.zapier.com/cli_docs/docs#using-http-middleware
 */

const includeBearerToken = (request, z, bundle) => {
  if (bundle.authData.access_token) {
    request.headers.Authorization = `Bearer ${bundle.authData.access_token}`;
  }
  return request;
};

const checkForErrors = (response, z) => {
  // In some cases the lower levels of the code can provide better error
  // messages. If they set this flag in the request, then we should fall back
  // to them to handle the errors.
  if (response.request.disableMiddlewareErrorChecking) {
    return response;
  }

  const responseHttpStatusCode = response.status;
  // Don't do any error message checking if we get a response in the 200 status code range
  if (responseHttpStatusCode >= 200 && responseHttpStatusCode < 299) {
    return response;
  }

  // In other cases, we may want to rely on JSON error message parsing from below
  // but also provide a more descriptive prefix.
  const { prefixErrorMessageWith } = response.request;

  let responseJson;
  try {
    responseJson = z.JSON.parse(response.content);
    // Don't throw any new Errors in here
  } catch (ex) {
    // Ignore it, we'll handle it below
  }

  if (responseJson && responseJson.error) {
    // Due to the way Microsoft APIs work - we need to pass in what scopes we want every time
    // we refresh an access token. As we add new features, we may want to add more scopes. If
    // we added the scopes and migrated users, their Zaps would break because they haven't
    // approved those new scopes. If they want to use the new features, they'll need to
    // reconnect their account. Let's provide a valuable and clear error message for them to
    // make that clear.
    if (responseJson.error.code === 'ErrorAccessDenied') {
      throw new Error(
        `${prefixErrorMessageWith}: This feature requires new permissions from your Exchange account. Please reconnect your account to take advantage of it.`
      );
    } else if (responseJson.error.code === 'ErrorInvalidIdMalformed') {
      throw new z.errors.HaltedError(
        `${prefixErrorMessageWith}: One of the fields you entered has an invalid id.`
      );
    } else if (
      responseHttpStatusCode === 413 &&
      responseJson.error.code === 'BadRequest' &&
      responseJson.error.message === 'Maximum request length exceeded.'
    ) {
      throw new z.errors.HaltedError(
        `${prefixErrorMessageWith}: Attached files must be less than 4MB.`
      );
    } else {
      // Ideally most APIs will return an error with this structure.
      throw new Error(
        `${prefixErrorMessageWith}: ${responseJson.error.message}`
      );
    }
  }

  // If we end up here guess we don't really have any idea what happened so we'll
  // just return the generic content.
  throw new Error(
    `${prefixErrorMessageWith}. Error code ${responseHttpStatusCode}: ${
      response.content
    }`
  );
};

module.exports = {
  includeBearerToken,
  checkForErrors,
};
