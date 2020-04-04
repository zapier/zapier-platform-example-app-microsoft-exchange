const { triggerNewContact } = require('zapier-platform-common-microsoft/triggers/new_contact');

const sample = require('../samples/contact');

module.exports = {
  key: 'new_contact',
  noun: 'Contact',
  display: {
    label: 'New Contact',
    description: 'Triggers when a new contact is added to your account',
  },

  operation: {
    inputFields: [
      {
        key: 'contactFolderId',
        label: 'Contact Folder',
        helpText: 'To trigger only on new contacts in a certain folder, select a folder here. If empty, we will trigger on new contacts in the default Contacts folder.',
        dynamic: 'list_contact_folders.id.displayName',
      },
    ],
    perform: triggerNewContact,

    sample,
  },
};
