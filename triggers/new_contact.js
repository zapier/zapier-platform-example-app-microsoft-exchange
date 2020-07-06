const { cleanContactEntries, getContactRequestURL } = require('../helpers');

const triggerNewContact = async (z, bundle) => {
  const response = await z.request({
    url: getContactRequestURL(bundle),
    params: {
      $orderby: 'createdDateTime desc',
      $top: 50,
    },
    prefixErrorMessageWith: 'Unable to retrieve the list of contacts',
  });
  const contacts = z.JSON.parse(response.content).value;
  return contacts.map(cleanContactEntries);
};

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
        helpText:
          'To trigger only on new contacts in a certain folder, select a folder here. If empty, we will trigger on new contacts in the default Contacts folder.',
        dynamic: 'list_contact_folders.id.displayName',
      },
    ],
    perform: triggerNewContact,

    sample,
  },
};
