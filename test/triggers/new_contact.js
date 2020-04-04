const { API_BASE_URL } = require('zapier-platform-common-microsoft/constants');
const { expect } = require('chai');
const zapier = require('zapier-platform-core');
const nock = require('nock');
const App = require('../../index');

const appTester = zapier.createAppTester(App);
const getBundle = () => ({
  inputData: {},
});

describe('Microsoft Exchange App', () => {
  it('declares a new contact trigger', () => expect(App.triggers.new_contact).to.exist);

  describe('new contact trigger', () => {
    describe('given no contact folder is selected and Exchange returns a valid list of contacts', () => {
      const getContactsResponse = require('../fixtures/responses/list_contacts_response'); // eslint-disable-line global-require

      let result;
      before(async () => {
        nock(API_BASE_URL)
          .get('/me/contacts')
          .query({ $orderby: 'createdDateTime desc', $top: '50' })
          .reply(200, getContactsResponse);

        // when the user tries to get a list of contacts
        result = await appTester(App.triggers.new_contact.operation.perform);
      });

      it('returns the expected contacts', () => {
        expect(result.length).to.eql(10);
        expect(result[0].id).to.eql(getContactsResponse.value[0].id);
        expect(result[1].id).to.eql(getContactsResponse.value[1].id);
        expect(result[0].emailAddresses_1).to.eql('zap.zaplar@zapier.ninja');
      });

      it('removes unnecessary attributes from the returned object', () => {
        expect(result).to.not.contain.any.item.with.property('@odata.etag');
      });
    });

    describe('given a contact folder is selected and Exchange returns a valid list of contacts', () => {
      const getContactsInFolderResponse = require('../fixtures/responses/list_contacts_response'); // eslint-disable-line global-require

      let result;
      before(async () => {
        const bundle = getBundle();
        bundle.inputData.contactFolderId = 'someContactFolderId';
        nock(API_BASE_URL)
          .get(`/me/contactFolders/${bundle.inputData.contactFolderId}/contacts`)
          .query({ $orderby: 'createdDateTime desc', $top: '50' })
          .reply(200, getContactsInFolderResponse);

        // when the user tries to get a list of contacts
        result = await appTester(App.triggers.new_contact.operation.perform, bundle);
      });

      it('returns the expected contacts', () => {
        expect(result.length).to.eql(10);
        expect(result[0].id).to.eql(getContactsInFolderResponse.value[0].id);
        expect(result[1].id).to.eql(getContactsInFolderResponse.value[1].id);
        expect(result[0].emailAddresses_1).to.eql('zap.zaplar@zapier.ninja');
        expect(result).to.be.an('array').to.all.include.deep.property('parentFolderId', 'someContactFolderId');
      });

      it('removes unnecessary attributes from the returned object', () => {
        expect(result).to.not.contain.any.item.with.property('@odata.etag');
      });
    });

    describe('given Exchange returns no contact', () => {
      let result;
      before(async () => {
        nock(API_BASE_URL)
          .get('/me/contacts')
          .query({ $orderby: 'createdDateTime desc', $top: '50' })
          .reply(200, {
            value: [],
          });

        // when the user tries to get a list of contacts
        result = await appTester(App.triggers.new_contact.operation.perform);
      });

      it('returns an empty list with no contacts', () => {
        expect(result.length).to.eql(0);
      });
    });

    describe('given Exchange returns an invalid response', () => {
      before(() => {
        nock(API_BASE_URL)
          .get('/me/contacts')
          .query({ $orderby: 'createdDateTime desc', $top: '50' })
          .reply(500, { error: { message: 'some random error' } });
      });

      it('returns a descriptive message', async () => {
        await expect(
          appTester(App.triggers.new_contact.operation.perform),
        ).to.be.rejectedWith('Unable to retrieve the list of contacts: some random error');
      });
    });
  });
});
