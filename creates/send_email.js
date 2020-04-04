const { sendEmail } = require('zapier-platform-common-microsoft/creates/send_email');

const sample = require('../samples/simple_email');

module.exports = {
  key: 'send_email',
  noun: 'Email',

  display: {
    label: 'Send Email',
    description: 'Send an email from your Exchange account.',
    important: true,
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
    perform: sendEmail,

    sample,
  },
};
