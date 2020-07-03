const authentication = require('./authentication');
const middleware = require('./middleware');

const FindContact = require('./searches/find_contact');
const NewContactTrigger = require('./triggers/new_contact');
const CreateContact = require('./creates/create_contact');
const ListContactFoldersTrigger = require('./triggers/list_contact_folders');

const App = {
  version: require('./package.json').version, // eslint-disable-line global-require
  platformVersion: require('zapier-platform-core').version, // eslint-disable-line global-require

  authentication,

  beforeRequest: [middleware.includeBearerToken],

  afterResponse: [middleware.checkForErrors],

  triggers: {
    [NewContactTrigger.key]: NewContactTrigger,
    [ListContactFoldersTrigger.key]: ListContactFoldersTrigger,
  },

  searches: {
    [FindContact.key]: FindContact,
  },

  creates: {
    [CreateContact.key]: CreateContact,
  },

  searchOrCreates: {
    [FindContact.key]: {
      key: FindContact.key,
      display: {
        label: 'Find or Create Contact',
        description: 'this is the description.', // this is ignored
      },
      search: FindContact.key,
      create: CreateContact.key,
    },
  },
};

module.exports = App;
