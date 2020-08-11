# Microsoft Exchange

[![Build Status](https://travis-ci.com/zapier/zapier-platform-cli-apps.svg?token=usX6G3kJjLapz4YDeEzM&branch=master)](https://travis-ci.com/zapier/zapier-platform-cli-apps)
[![code style: prettier](https://img.shields.io/badge/code_style-prettier-ff69b4.svg?style=flat-square)](https://github.com/prettier/prettier)

This is an example CLI App for the Zapier Developer Platform. It uses the Microsoft Exchange API, showing what a fully-featured integration looks like. It follows the best practices for building an app that we follow here at Zapier.

What it demonstrates:

- Polling Trigger
- Search action
- Basic action
- Dynamic dropdowns
- Middleware (examples of sensible error-handling)
- OAuth2

## About the Zapier Platform

Zapier is an app automation platform where over 2 million people connect apps into customized workflows—what we call Zaps. A Zapier integration can connect your app’s API to the Zapier platform and let it integrate with 1,300+ other popular apps to watch for new or updated data, find existing data, or create and update data.

A Zapier integration for your app lets your users build workflows like that and more with your app.

If you're new to building Zapier integrations, here are some resources to get you started: 

- [Getting Started With Zapier Platform CLI](https://platform.zapier.com/cli_tutorials/getting-started)
- [Zapier Platform CLI Documentation](https://platform.zapier.com/cli_docs/docs)

## Developing

### Built With

- [Node](https://nodejs.org/en/) `8.10.0` (This matches with the AWS Lamba Node Verison)
- [Chai](https://www.chaijs.com/)

### Prerequisites

You'll need to have the Zapier Platform CLI installed if you haven't already:

```shell
npm install -g zapier-platform-cli
```

### Setting up Dev

Run the following commands to get started locally:

```shell
git clone git@github.com:zapier/zapier-platform-cli-apps.git
cd zapier-platform-cli-apps/apps/microsoft-exchange
npm install
```

Run `zapier register 'My Example MS Exchange App'` to register the app with Zapier.

You'll need to register a new app in Microsoft Azure, which you can do here: https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade

The app will have a Client ID and Client Secret, which you can set by following the instructions here: https://platform.zapier.com/cli_docs/docs#defining-environment-variables

## Versioning

We should use [SemVer](http://semver.org/) for versioning. For the versions available, see the
[CHANGELOG](CHANGELOG.md).

## Tests

Right now all we have are unit tests. You can run them with the command `zapier test`.

## Style guide

We use [ESLint](https://eslint.org/) and a common set of rules developed by AirBnB with a couple of tweaks. Linting
should occur automatically whenever you make a commit (see [package.json](package.json) for the exact things being run).

With that being said, you can run the following command to run the lint check against all of your files:

```shell
npm run lint
```

Or you can run this to format all of your files properly:

```shell
npm run format
```

## Api References

https://docs.microsoft.com/en-us/graph/use-the-api

## Contributing

We are not accepting contributions to this repo. This is just an example integration that illustrates a more comprehensive example of a Zapier integration built using the Zapier CLI.

This repo isn't actively maintained.
