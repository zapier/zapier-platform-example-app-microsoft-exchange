/**
 * This trigger pulls choices for the "Contact Folder" field. See the comments in `new_contact.js` for more details.
 * You'll notice in the exported definition that `display.hidden` is set to `true`.
 */

const { API_BASE_URL } = require('../constants');

const triggerListContactFolders = async z => {
  const response = await z.request({
    url: `${API_BASE_URL}/me/contactFolders`,
    prefixErrorMessageWith: 'Unable to retrieve the list of folders',
  });

  return z.JSON.parse(response.content).value;
};

const sample = require('../samples/contact_folder');

module.exports = {
  key: 'list_contact_folders',
  noun: 'List Contact Folders',

  display: {
    label: "Get a list of the contact folders in the users' account",
    description: 'Used to power a folder dropdown for contacts.',
    hidden: true,
  },

  operation: {
    inputFields: [
      {
        key: 'contactFolderId',
        label: 'Contact Folder',
        helpText: 'If empty, the default Contacts folder will be used.',
        dynamic: 'list_contact_folders.id.displayName',
      },
    ],
    perform: triggerListContactFolders,

    sample,
  },
};
