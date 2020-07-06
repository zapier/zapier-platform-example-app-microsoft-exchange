const { API_BASE_URL } = require('./constants');

/**
 * Takes bundle.inputData and extracts/formats the array of emails/name strings
 * into the format that Microsoft expects for emails. If the email array passed
 * in is undefined or empty, we'll return an empty array (which
 * is acceptable for the Microsoft APIs).
 *
 * Example:
 * inputData = {
        givenName: 'firstName',
        surname: 'lastName',
        emailAddresses: [
          'emailOne@test.com',
          'emailTwo@testing.com',
          'emailThree@testing1.com'
        ],
    }
 *
 * output: [
        {"name": "firstName lastName", "address": "emailOne@test.com"},
        {"name": "firstName lastName", "address": "emailTwo@testing.com"},
        {"name": "firstName lastName", "address": "emailThree@testing1.com"}
    ]
 */
const formatContactEmails = inputData => {
  const { emailAddresses, givenName, surname } = inputData;
  if (emailAddresses) {
    const result = [];
    emailAddresses.forEach(email => {
      result.push({
        name: givenName + (surname ? ' ' + surname : ''),
        address: email,
      });
    });
    return result;
  }
  return [];
};

/**
 * Takes in an object within an array [{}] of address child keys
 * and modifies the keys into the format that Microsoft expects for addresses.
 * If the array passed in is undefined or empty, we'll return an empty object
 * (which is acceptable for the Microsoft APIs).
 *
 * Example:
 * input: [
 *     {
        businessAddress_street: '123 fake st',
        businessAddress_city: 'austin',
        businessAddress_state: 'tx',
        businessAddress_postalCode: '78701',
        businessAddress_countryOrRegion: 'us'
        },
      ]
 * output: {
      "street": "123 fake st",
      "city": "austin",
      "state": "tx",
      "postalCode": "78701"
      "countryOrRegion": "us"
    }
 */
const formatContactAddress = addressData => {
  let formattedAddress = {};
  if (addressData) {
    for (key in addressData[0]) {
      formattedAddress[key.split('_')[1]] = addressData[0][key];
    }
  }
  return formattedAddress;
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

module.exports = {
  formatContactEmails,
  formatContactAddress,
  cleanContactEntries,
  getContactRequestURL,
};
