const { createContact } = require('zapier-platform-common-microsoft/creates/create_contact');

const sample = require('../samples/contact');

module.exports = {
  key: 'create_contact',
  noun: 'Contact',

  display: {
    label: 'Create Contact',
    description: 'Creates a new contact.',
  },

  operation: {
    inputFields: [
      {
        key: 'contactFolderId',
        label: 'Contact Folder',
        helpText: 'If empty, a contact will be created in the default Contacts folder.',
        dynamic: 'list_contact_folders.id.displayName',
      },
      {
        key: 'givenName',
        label: 'First Name',
        required: true,
      },
      { key: 'surname', label: 'Last Name' },
      {
        key: 'emailAddresses',
        helpText: 'Microsoft allows a maximum of `THREE` email addresses when creating contacts.',
        list: true,
      },
      {
        key: 'businessPhones',
        helpText: 'Microsoft allows a maximum of `TWO` business phone numbers when creating contacts.',
        list: true,
      },
      {
        key: 'homePhones',
        helpText: 'Microsoft allows a maximum of `TWO` home phone numbers when creating contacts.',
        list: true,
      },
      { key: 'mobilePhone' },
      { key: 'jobTitle' },
      { key: 'companyName' },
      { key: 'department' },
      { key: 'businessHomePage', label: 'Business Website URL' },
      { key: 'fileAs', helpText: 'Optionally, specify a name to file the contact under.' },
      { key: 'personalNotes', type: 'text' },
      {
        key: 'businessAddress',
        children: [
          { key: 'businessAddress_street', label: 'Street' },
          { key: 'businessAddress_city', label: 'City' },
          { key: 'businessAddress_state', label: 'State' },
          { key: 'businessAddress_postalCode', label: 'Postal Code' },
          { key: 'businessAddress_countryOrRegion', label: 'Country or Region' },
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
    perform: createContact,

    sample,
  },
};
