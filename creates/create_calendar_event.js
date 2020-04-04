const {
  createCalendarEvent,
} = require('zapier-platform-common-microsoft/creates/create_calendar_event');

const sample = require('../samples/calendar_event');

module.exports = {
  key: 'create_calendar_event',
  noun: 'Event',

  display: {
    label: 'Create Event',
    description: 'Create an event in the calendar of your choice.',
    important: true,
  },

  operation: {
    inputFields: [
      {
        key: 'calendarId',
        label: 'Calendar',
        helpText: 'If empty, will default to your default calendar.',
        dynamic: 'list_calendars.id',
      },
      {
        key: 'subject',
        label: 'Subject',
        type: 'string',
        required: true,
      },
      {
        key: 'startTime',
        label: 'Start Date & Time',
        type: 'datetime',
        helpText:
          'Date and time of when this event starts. See this [help doc](https://zapier.com/help/modifying-dates-and-times/#adjusting-dates-and-times) for more information.',
        required: true,
      },
      {
        key: 'endTime',
        label: 'End Date & Time',
        type: 'datetime',
        helpText: 'Date and time of when this event ends.',
        required: true,
      },
      {
        key: 'isAllDay',
        label: 'All Day Event?',
        helpText:
          'If "true", we\'ll ignore the time entered in the start and end time fields. For example, `2019-03-25 6pm` would become `2019-03-25`.',
        type: 'boolean',
      },
      {
        key: 'body',
        label: 'Description',
        helpText: 'Allows HTML',
        type: 'text',
      },
      {
        key: 'showAs',
        label: 'Show me as Free or Busy',
        placeholder: 'Busy',
        choices: ['free', 'busy'],
      },
    ],
    perform: createCalendarEvent,

    sample,
  },
};
