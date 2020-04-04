const { findContact } = require('zapier-platform-common-microsoft/searches/find_contact');

const sample = require('../samples/contact');

module.exports = {
  key: 'find_contact',
  noun: 'Contact',
  display: {
    label: 'Find a Contact',
    description: 'Search for a contact.',
  },

  operation: {
    inputFields: [
      {
        key: 'contactFolderId',
        label: 'Contact Folder',
        helpText: 'If empty, will search for contacts in the default Contacts folder.',
        dynamic: 'list_contact_folders.id.displayName',
      },
      {
        key: 'givenName',
        label: 'First Name',
      },
      { key: 'surname', label: 'Last Name' },
      // setting this key as emailAddresses to match the Create Contact
      // action so fields don't get duplicated on the Find/Create action
      {
        key: 'emailAddresses', label: 'Email',
      },
    ],
    perform: findContact,

    sample,
  },
};
