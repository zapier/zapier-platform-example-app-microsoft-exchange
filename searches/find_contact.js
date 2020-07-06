const { cleanContactEntries, getContactRequestURL } = require('../helpers');

const findContact = async (z, bundle) => {
  const filters = [];
  const { emailAddresses, givenName, surname } = bundle.inputData;
  if (!emailAddresses && !givenName && !surname) {
    throw new Error(
      'Please enter a value in one of the search action input fields'
    );
  }
  if (emailAddresses) {
    filters.push(`emailAddresses/any(a:a/address eq '${emailAddresses}')`);
  }
  if (givenName) {
    filters.push(`givenName eq '${givenName}'`);
  }
  if (surname) {
    filters.push(`surname eq '${bundle.inputData.surname}'`);
  }
  const response = await z.request({
    url: getContactRequestURL(bundle),
    params: {
      $top: 50,
      $filter: filters.join(' and '),
    },
    prefixErrorMessageWith: 'Unable to retrieve the list of contacts',
  });
  const contacts = z.JSON.parse(response.content).value;
  return contacts.map(cleanContactEntries);
};

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
        helpText:
          'If empty, will search for contacts in the default Contacts folder.',
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
        key: 'emailAddresses',
        label: 'Email',
      },
    ],
    perform: findContact,

    sample,
  },
};
