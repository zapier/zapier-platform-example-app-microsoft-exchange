const { createDraftEmail } = require('zapier-platform-common-microsoft/creates/create_draft_email');
const sample = require('../samples/email');

module.exports = {
  key: 'create_draft_email',
  noun: 'Draft Email',

  display: {
    label: 'Create Draft Email',
    description: 'Creates a draft of an email that can then be reviewed and sent out.',
  },

  operation: {
    inputFields: [
      {
        key: 'recipients',
        label: 'To Email(s)',
        type: 'string',
        list: true,
        required: true,
      },
      {
        key: 'ccRecipients',
        label: 'CC Email(s)',
        type: 'string',
        list: true,
      },
      {
        key: 'bccRecipients',
        label: 'BCC Email(s)',
        type: 'string',
        list: true,
      },
      {
        key: 'subject',
        label: 'Subject',
        type: 'string',
        required: true,
      },
      {
        key: 'bodyFormat',
        label: 'Body Format',
        choices: ['Text', 'HTML'],
        required: true,
      },
      {
        key: 'body',
        label: 'Body',
        type: 'text',
        required: true,
      },
      {
        key: 'file',
        label: 'Attachment',
        type: 'file',
        helpText:
          'A file to be attached to the email. Can be an actual file or a public URL which will be downloaded and attached. '
          + 'Plain text content will converted into a txt file and attached. Please note the maximum size for an attachment is 4MB.',
      },
    ],
    perform: createDraftEmail,

    sample,
  },
};
