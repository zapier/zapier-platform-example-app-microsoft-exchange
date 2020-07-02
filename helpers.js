const { API_BASE_URL } = require('./constants');

const moment = require('moment');
const contentDisposition = require('content-disposition');
// TODO: investigate what content-disposition is

/**
 * Takes an array of strings of emails and translates it into the
 * format that Microsoft expects for emails. If the array passed
 * in is undefined or empty, we'll return an empty array (which
 * is acceptable for the Microsoft APIs).
 *
 * Example:
 * input: [foo@abc.com, bar@def.com]
 *
 * output: [
        {
            "emailAddress":{
                "address":"foo@abc.com"
            },
            "emailAddress":{
                "address":"bar@def.com"
            },
        }
    ]
 */
const formatEmails = recipients => {
  if (recipients === undefined || recipients.length === 0) {
    return [];
  }

  return recipients.map(recipient => ({
    emailAddress: { address: recipient },
  }));
};

/**
 * Cleans up the start / end times of an entry to format them in Zapier
 * standard time.
 *
 * Also removes unnecessary and confusing odata tags.
 */
const cleanEventEntries = entry => {
  const startDateTime = moment.utc(entry.start.dateTime);
  const endDateTime = moment.utc(entry.end.dateTime);

  delete entry.start;
  delete entry.end;

  delete entry['@odata.etag'];
  delete entry['@odata.context'];

  entry.startDateTime = startDateTime.format();
  entry.endDateTime = endDateTime.format();

  return entry;
};

/**
 * Returns the expected event requestUrl depending on what parameters the
 * user has passed in to Zapier.
 */
const getEventRequestURL = bundle => {
  let requestUrl = `${API_BASE_URL}/me/calendar/events`;

  if (bundle && bundle.inputData && bundle.inputData.calendarId) {
    requestUrl = `${API_BASE_URL}/me/calendars/${
      bundle.inputData.calendarId
    }/events`;
  }

  return requestUrl;
};

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
 * We need to check if the user specifies any address data when we try
 * and make an update call to decide if we should include the key or
 * not. If the user does not specify anything and we include the key,
 * we will wipe out all of their existing address data. We only want to
 * send over address updates if the user specifies anything.
 */
const checkIfAnyAddressDataIsSpecified = addressData => {
  let addressDataIsSpecified = false;

  if (addressData) {
    for (key in addressData[0]) {
      if (addressData[0][key]) {
        addressDataIsSpecified = true;
      }
    }
  }

  return addressDataIsSpecified;
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

/**
 * Takes in a file object, url, or plain text and converts it to a base64 encoded string
 * which is required when adding attachments to outbound emails.
 */
const formatEmailAttachment = async (z, file) => {
  if (!file) {
    return [];
  }
  const request = await z.request({
    url: file,
    raw: true,
  });
  // .buffer() returns a promise which needs to be base64 encoded to a string
  const buffer = await request.buffer();
  const base64String = Buffer.from(buffer).toString('base64');

  // if the content type is not included in the file name it could be marked
  // as a "suspicious" attachment. set the content-type in the request
  const fileType = request.headers.get('content-type');

  // use the content-disposition (which _should_ always exist) to get the filename.
  // if the content-disposition does not exist, default to 'unknown.unknown'
  const disposition = request.headers.get('content-disposition');
  let filename;
  if (disposition) {
    filename = contentDisposition.parse(disposition).parameters.filename;
  } else {
    filename = 'unknown.unknown';
  }

  return [
    {
      '@odata.type': '#Microsoft.graph.fileAttachment',
      name: filename,
      contentBytes: base64String,
      contentType: fileType,
    },
  ];
};

/**
 * Returns the expected requestUrl for updating contacts depending on what
 * parameters the user has passed in to Zapier.
 */
const updateContactRequestURL = bundle => {
  if (!bundle || !bundle.inputData || !bundle.inputData.contact) {
    // This should never happen but putting it here just in case
    throw new Error('Please specify a contact to update.');
  }

  let requestUrl = `${API_BASE_URL}/me/contacts/${bundle.inputData.contact}`;

  if (bundle && bundle.inputData && bundle.inputData.contactFolderId) {
    requestUrl = `${API_BASE_URL}/me/contactFolders/${
      bundle.inputData.contactFolderId
    }/contacts/${bundle.inputData.contact}`;
  }

  return requestUrl;
};

module.exports = {
  formatEmails,
  cleanEventEntries,
  getEventRequestURL,
  formatContactEmails,
  formatContactAddress,
  cleanContactEntries,
  getContactRequestURL,
  formatEmailAttachment,
  updateContactRequestURL,
  checkIfAnyAddressDataIsSpecified,
};
