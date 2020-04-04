const {
  stashEmailFunction,
} = require('zapier-platform-common-microsoft/utils');

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

  hydrators: {
    stashEmailFunction,
  },

  triggers: {
    [ListCalendarsTrigger.key]: ListCalendarsTrigger,
    [NewCalendarEventTrigger.key]: NewCalendarEventTrigger,
    [NewEmailTrigger.key]: NewEmailTrigger,
    [CalendarEventStartTrigger.key]: CalendarEventStartTrigger,
    [UpdatedCalendarEventTrigger.key]: UpdatedCalendarEventTrigger,
    [NewContactTrigger.key]: NewContactTrigger,
    [ListContactFoldersTrigger.key]: ListContactFoldersTrigger,
  },

  searches: {
    [FindContact.key]: FindContact,
  },

  creates: {
    [CreateCalendarEvent.key]: CreateCalendarEvent,
    [CreateDraftEmail.key]: CreateDraftEmail,
    [SendEmail.key]: SendEmail,
    [CreateContact.key]: CreateContact,
    [UpdateContact.key]: UpdateContact,
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
