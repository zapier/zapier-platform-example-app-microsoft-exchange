// This is where the core definition of your app is set
// Relevant Docs: https://platform.zapier.com/cli_docs/docs#local-project-structure

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

  beforeRequest: [middleware.includeBearerToken], // Docs for middleware: https://platform.zapier.com/cli_docs/docs#using-http-middleware

  afterResponse: [middleware.checkForErrors], // Docs for middleware: https://platform.zapier.com/cli_docs/docs#using-http-middleware

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
