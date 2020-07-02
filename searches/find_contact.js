/**
 * Returns the expected contact requestUrl depending on what parameters the
 * user has passed in to Zapier.
 */
const getContactRequestURL = bundle => {
  let requestUrl = `${API_BASE_URL}/me/contacts`;

  if (bundle && bundle.inputData && bundle.inputData.contactFolderId) {
    requestUrl = `${API_BASE_URL}/me/contactFolders/${
      bundle.inputData.contactFolderId
    }/contacts`;
  }

  return requestUrl;
};

/**
 * Removes unnecessary and confusing odata tags.
 *
 * Also, it add properties for each email and business/home phone number
 * as so users can individually map them
 *
 * Example:
 * input: {
        businessPhones: ['555-555-5555', '555-555-5554'],
        homePhones: ['123-555-5553', '123-555-5552'],
        emailAddresses: [
          {name: 'test contact', address: 'emailOne@test.com'},
          {name: 'test contact', address: 'emailTwo@test.com'},
          {name: 'test contact', address: 'emailThree@test.com'},
        ],
    }
 *
 * output: {
 *      "businessPhones": ["555-555-5555", "555-555-5554"],
        "homePhones": ["123-555-5553", "123-555-5552"],
        "emailAddresses": [
          {"name": "test contact", "address": "emailOne@test.com"},
          {"name": "test contact", "address": "emailTwo@test.com"},
          {"name": "test contact", "address": "emailThree@test.com"},
        ],
        "businessPhones_1": "555-555-5555",
        "businessPhones_2": "555-555-5554",
        "homePhones_1": "123-555-5553",
        "homePhones_2": "123-555-5552",
        "emailAddresses_1": "emailOne@test.com",
        "emailAddresses_2": "emailTwo@test.com",
        "emailAddresses_3": "emailThree@test.com"
      }
 */
const cleanContactEntries = entry => {
  delete entry['@odata.etag'];
  delete entry['@odata.context'];

  const keysToUnwrap = ['businessPhones', 'homePhones', 'emailAddresses'];

  keysToUnwrap.forEach(key => {
    entry[key].forEach((number, index) => {
      if (key !== 'emailAddresses') {
        entry[`${key}_${index + 1}`] = number;
      } else {
        entry[key].forEach((emailObject, index) => {
          entry[`${key}_${index + 1}`] = emailObject.address;
        });
      }
    });
  });

  return entry;
};

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
