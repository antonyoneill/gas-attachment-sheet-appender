# GAS Attachment Sheet Appender

**_Unsupported - use at your own risk_**

This was thrown together to meet a business requirement for an online retailer.

It is designed to be executed within Google App Scripts.

It meets the following criteria:

1. Scan for new emails within a given Gmail Label
2. For each .xls attachment:
   1. Convert it to a Google Spreadsheet within a temporary folder `temporaryDriveFolderId`
   2. Append the data to a number of existing Spreadsheets `driveTargetSheetIds`
   3. Optionally: Convert to CSV and copy out to another folder `config.additionalCSVExport[].driveFolderId`
   4. Optionally: Execute a callback function with the csv blob

## Installation

1. Go to Google Drive
2. Create a new 'Google App Script' file
3. Copy `src/index.js`
4. [Enable Drive Advanced Google Services](https://developers.google.com/apps-script/guides/services/advanced#enabling_advanced_services)
5. Set config properties to suit your requirements
6. Set up triggers to run this automatically

## Uploading the CSV to AWS S3

There exists another library that provides this functionality - provided kindly by https://github.com/eschultink/S3-for-Google-Apps-Script

If you'd like to copy the CSV up to S3 then you can use a callback with the AWS client. Something like:

```javascript
const callback = csvString => {
  const awsAccessKeyId = "MY-AWS-ACCESS-KEY-ID";
  const awsSecretKey = "MY-AWS-SECRET-KEY";
  var s3 = getInstance(awsAccessKeyId, awsSecretKey);

  s3.putObject(
    "MY-BUCKET-NAME",
    new Date().toISOString() + ".csv",
    csvString,
    {
      logRequests: true,
    }
  );
};
```

## Testing

_Forgive me for I have sinned._

As much as I preach TDD and the importance of testing, given this was a 2-3 hour project with a tight deadline, and the time taken to learn the new GAS / S3 APIs there are no tests.

One day I will write some, and likely rewrite the interface because it's a bit all over the place.

## Architecture

Well, you can see that it all exists within a single JavaScript file. That's not great.

I have made an effort to separate out the concerns and have logical functions.

## Support

**There is absolutely no support provided by myself for your usage of this script.**

As described above, and by looking at the repo, you will see there are no tests written for this.

I have tested it manually myself, following golden paths, and a limited number of edge cases, but I can not guarantee the quality of the code written.

## License

[MIT](./LICENSE)

