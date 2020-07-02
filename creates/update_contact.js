const {
  formatContactEmails,
  formatContactAddress,
  cleanContactEntries,
  updateContactRequestURL,
  checkIfAnyAddressDataIsSpecified,
} = require('../helpers');

const updateContact = async (z, bundle) => {
  const {
    givenName,
    surname,
    businessPhones,
    homePhones,
    mobilePhone,
    jobTitle,
    companyName,
    department,
    businessHomePage,
    personalNotes,
  } = bundle.inputData;

  const body = {
    givenName,
    surname,
    emailAddresses: formatContactEmails(bundle.inputData),
    businessPhones,
    homePhones,
    mobilePhone,
    jobTitle,
    companyName,
    department,
    businessHomePage,
    personalNotes,
    businessAddress: formatContactAddress(bundle.inputData.businessAddress),
    homeAddress: formatContactAddress(bundle.inputData.homeAddress),
    otherAddress: formatContactAddress(bundle.inputData.otherAddress),
  };

  // When a user does not specify something in these fields in Zapier, we
  // end up defaulting them to [] - which Microsoft thinks means "delete all
  // existing data for these fields". In reality, we don't want to touch fields
  // that the user does not specify so let's remove any keys from the body
  // that the user hasn't specified.
  if (
    !bundle.inputData.emailAddresses ||
    bundle.inputData.emailAddresses.length === 0
  ) {
    delete body.emailAddresses;
  }

  if (!checkIfAnyAddressDataIsSpecified(bundle.inputData.businessAddress)) {
    delete body.businessAddress;
  }

  if (!checkIfAnyAddressDataIsSpecified(bundle.inputData.homeAddress)) {
    delete body.homeAddress;
  }

  if (!checkIfAnyAddressDataIsSpecified(bundle.inputData.otherAddress)) {
    delete body.otherAddress;
  }

  // Existing properties that are not included in the request body will maintain
  // their previous values or be recalculated based on changes to other property
  // values. Passing in "null" for a field does not cause it to change.
  //
  // Nested values, such as addresses, need to include all of the fields, however.
  // If a user currently has a city / state specified and they make an update
  // and only specify the Zip code - the city and state will be removed.
  const response = await z.request({
    method: 'PATCH',
    url: updateContactRequestURL(bundle),
    json: body,
    prefixErrorMessageWith: 'Unable to update the specified contact',
  });

  const parsedJson = z.JSON.parse(response.content);

  return cleanContactEntries(parsedJson);
};

const sample = require('../samples/contact');

module.exports = {
  key: 'update_contact',
  noun: 'Contact',

  display: {
    label: 'Update Contact',
    description: 'Updates a contact.',
  },

  operation: {
    inputFields: [
      {
        key: 'contactFolderId',
        label: 'Contact Folder',
        helpText:
          'If empty, the dropdown below will only show contacts in your top level contactFolder.',
        dynamic: 'list_contact_folders.id.displayName',
        altersDynamicFields: true,
      },
      {
        key: 'contact',
        label: 'Contact',
        dynamic: 'new_contact.id.displayName',
        search: 'find_contact.id.displayName',
        required: true,
      },
      {
        key: 'givenName',
        label: 'First Name',
      },
      { key: 'surname', label: 'Last Name' },
      {
        key: 'emailAddresses',
        helpText:
          'Microsoft allows a maximum of `THREE` email addresses when updating contacts.',
        list: true,
      },
      {
        key: 'businessPhones',
        helpText:
          'Microsoft allows a maximum of `TWO` business phone numbers when updating contacts.',
        list: true,
      },
      {
        key: 'homePhones',
        helpText:
          'Microsoft allows a maximum of `TWO` home phone numbers when updating contacts.',
        list: true,
      },
      { key: 'mobilePhone' },
      { key: 'jobTitle' },
      { key: 'companyName' },
      { key: 'department' },
      { key: 'businessHomePage', label: 'Business Website URL' },
      { key: 'personalNotes', type: 'text' },
      {
        key: 'businessAddress',
        children: [
          { key: 'businessAddress_street', label: 'Street' },
          { key: 'businessAddress_city', label: 'City' },
          { key: 'businessAddress_state', label: 'State' },
          { key: 'businessAddress_postalCode', label: 'Postal Code' },
          {
            key: 'businessAddress_countryOrRegion',
            label: 'Country or Region',
          },
        ],
      },
      {
        key: 'homeAddress',
        children: [
          { key: 'homeAddress_street', label: 'Street' },
          { key: 'homeAddress_city', label: 'City' },
          { key: 'homeAddress_state', label: 'State' },
          { key: 'homeAddress_postalCode', label: 'Postal Code' },
          { key: 'homeAddress_countryOrRegion', label: 'Country or Region' },
        ],
      },
      {
        key: 'otherAddress',
        children: [
          { key: 'otherAddress_street', label: 'Street' },
          { key: 'otherAddress_city', label: 'City' },
          { key: 'otherAddress_state', label: 'State' },
          { key: 'otherAddress_postalCode', label: 'Postal Code' },
          { key: 'otherAddress_countryOrRegion', label: 'Country or Region' },
        ],
      },
    ],
    perform: updateContact,

    sample,
  },
};
