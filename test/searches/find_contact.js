const { API_BASE_URL } = require('../../constants');
const { expect } = require('chai');
const zapier = require('zapier-platform-core');
const nock = require('nock');
const App = require('../../index');

const appTester = zapier.createAppTester(App);
const getBundle = () => ({
  inputData: {},
});

describe('Microsoft Exchange App', () => {
  it('declares a Find Contact search', () =>
    expect(App.searches.find_contact).to.exist);
  const findContactsResponse = require('../fixtures/responses/list_contacts_response'); // eslint-disable-line global-require

  describe('Find Contact search', () => {
    describe('given an email in the search field with no contact folder selected and a normal response', () => {
      let result;
      before(async () => {
        const bundle = getBundle();
        bundle.inputData.emailAddresses = 'zap.zaplar@zapier.ninja';
        const expectedFilter = `emailAddresses/any(a:a/address eq '${
          bundle.inputData.emailAddresses
        }')`;
        nock(API_BASE_URL)
          .get('/me/contacts')
          .query({ $top: '50', $filter: expectedFilter })
          .reply(200, findContactsResponse);

        result = await appTester(
          App.searches.find_contact.operation.perform,
          bundle
        );
      });

      it('returns the expected contact', () => {
        expect(result[0].id).to.eql(findContactsResponse.value[0].id);
        expect(result[0].emailAddresses_1).to.eql('zap.zaplar@zapier.ninja');
      });

      it('removes unnecessary attributes from the returned object', () => {
        expect(result).to.not.contain.any.item.with.property('@odata.etag');
      });
    });

    describe('given an email in the search field with a contact folder selected and a normal response', () => {
      let result;
      before(async () => {
        const bundle = getBundle();
        bundle.inputData.emailAddresses = 'zap.zaplar@zapier.ninja';
        bundle.inputData.contactFolderId = 'someContactFolderId';
        const expectedFilter = `emailAddresses/any(a:a/address eq '${
          bundle.inputData.emailAddresses
        }')`;
        nock(API_BASE_URL)
          .get(
            `/me/contactFolders/${bundle.inputData.contactFolderId}/contacts`
          )
          .query({ $top: '50', $filter: expectedFilter })
          .reply(200, findContactsResponse);

        result = await appTester(
          App.searches.find_contact.operation.perform,
          bundle
        );
      });

      it('returns the expected contact', () => {
        expect(result[0].id).to.eql(findContactsResponse.value[0].id);
        expect(result[0].emailAddresses_1).to.eql('zap.zaplar@zapier.ninja');
        expect(result[0].parentFolderId).to.eql('someContactFolderId');
      });

      it('removes unnecessary attributes from the returned object', () => {
        expect(result).to.not.contain.any.item.with.property('@odata.etag');
      });
    });

    describe('given a first/last name in the search field and a normal response', () => {
      let result;
      before(async () => {
        const bundle = getBundle();
        bundle.inputData = { givenName: 'Zap', surname: 'Zaplar' };
        const expectedFilter = `givenName eq '${
          bundle.inputData.givenName
        }' and surname eq '${bundle.inputData.surname}'`;
        nock(API_BASE_URL)
          .get('/me/contacts')
          .query({ $top: '50', $filter: expectedFilter })
          .reply(200, findContactsResponse);

        result = await appTester(
          App.searches.find_contact.operation.perform,
          bundle
        );
      });

      it('returns the expected contact', () => {
        expect(result[0].id).to.eql(findContactsResponse.value[0].id);
      });

      it('removes unnecessary attributes from the returned object', () => {
        expect(result).to.not.contain.any.item.with.property('@odata.etag');
      });
    });

    describe('given a search that has no match', () => {
      let result;
      before(async () => {
        const bundle = getBundle();
        bundle.inputData.emailAddresses = 'zip.ziplar@zapier.ninja';
        const expectedFilter = `emailAddresses/any(a:a/address eq '${
          bundle.inputData.emailAddresses
        }')`;
        nock(API_BASE_URL)
          .get('/me/contacts')
          .query({ $top: '50', $filter: expectedFilter })
          .reply(200, { value: [] });

        result = await appTester(
          App.searches.find_contact.operation.perform,
          bundle
        );
      });

      it('returns an empty array', () => {
        expect(result).to.eql([]);
      });
    });

    describe('given no search fields are filled out', () => {
      it('throws an error telling you to have some inputData please', async () => {
        await expect(
          appTester(App.searches.find_contact.operation.perform, getBundle())
        ).to.be.rejectedWith(
          'Please enter a value in one of the search action input fields'
        );
      });
    });

    describe('given Exchange returns an invalid response', () => {
      let bundle;
      before(() => {
        bundle = getBundle();
        bundle.inputData.emailAddresses = 'zap.zaplar@zapier.ninja';
        const expectedFilter = `emailAddresses/any(a:a/address eq '${
          bundle.inputData.emailAddresses
        }')`;
        nock(API_BASE_URL)
          .get('/me/contacts')
          .query({ $top: '50', $filter: expectedFilter })
          .reply(500, { error: { message: 'some random error' } });
      });

      it('returns a descriptive message', async () => {
        await expect(
          appTester(App.searches.find_contact.operation.perform, bundle)
        ).to.be.rejectedWith(
          'Unable to retrieve the list of contacts: some random error'
        );
      });
    });
  });
});
