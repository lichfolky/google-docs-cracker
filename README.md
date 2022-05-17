# Google Docs Cracker

Experiments with google cloud api and python.  
The `docs-gen.py script` create a doc in a new folder, for each row of a sheet,
using another document as template and substituting keywords with the row contents.

## Setup

### Install libs

```
  pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
```

### Setup a Google Cloud Platform project

- Create a project in [https://console.cloud.google.com/]
- Enable the APIs from **APIs & Services** (Google Drive API, Google Docs API, Google Sheets API )
- Create credentials (OAuth 2.0 Client IDs) from **APIs & Services** > Credentials
- Download json credentials file
- Use it to build the services

### Use the apis

- If you need to use a doc or a sheet, copy its id from the preview url

[Google sheets](https://developers.google.com/sheets/api/guides/concepts)  
[Google docs](https://developers.google.com/docs/api/how-tos/overview)  
[Google drive](https://developers.google.com/drive/api/guides/about-sdk)
