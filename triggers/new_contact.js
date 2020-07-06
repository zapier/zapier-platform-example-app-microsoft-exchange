const { cleanContactEntries, getContactRequestURL } = require('../helpers');

/**
 * This function will return a list of contacts from the API. The list is then passed
 * to a helper function, cleanContactEntries, to ensure the data is formatted in a way that allows
 * deduplication to happen properly. These docs go into greater detail:
 *
 * https://zapier.com/developer/documentation/v2/deduplication/
 * https://platform.zapier.com/cli_docs/docs#how-does-deduplication-work
 *
 */
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
        // A dynamic dropdown pulls a list of choices for the user to choose from—here's what it looks like in the UI: https://cdn.zappy.app/981bf505a5fb839b317cedb3548d050d.png
        // It uses the "list_contact_folders" trigger, which is a hidden trigger that serves to feed data to these dropdowns on the frontend.
        // Relevant docs: https://platform.zapier.com/cli_docs/docs#dynamic-dropdowns
        dynamic: 'list_contact_folders.id.displayName',
      },
    ],
    perform: triggerNewContact,

    sample,
  },
};
