const {
  updateContact,
} = require('zapier-platform-common-microsoft/creates/update_contact');

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
