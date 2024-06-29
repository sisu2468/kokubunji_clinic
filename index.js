const express = require('express');
const { google } = require('googleapis');
const bodyParser = require('body-parser');
const { JWT } = require('google-auth-library');

const app = express();
app.use(bodyParser.json()); // Middleware to parse JSON bodies

app.get('/data', async (req, res) => {
  const GOOGLE_SERVICE_ACCOUNT_EMAIL = "docusign@admin-310703.iam.gserviceaccount.com";
  const GOOGLE_PRIVATE_KEY = "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQCq9yvoxf9zZcHP\nlBWgb5N6dmuZcreS2w5q/0Il6sVIJfqDjflYic63ZWTAQ2O2uBvZJ4DFYT1QBcTW\nypzb0RVz0rnLc1NIrsEisTspTRbT7eLSyZUmLBPdVkCD6e+S/ZQWho/kLL3het2m\nC2rwTauT/T7KBgVyHyfr5AN0XvHFf98grevV1eSf3Aylg7mIKTti7DUL3NbOlOlJ\nGkPr3JoDrFfTPUREQO5gYiyQWif5lvWoPQhDjYIBQ2WXtlOtS+qpjUHmmyNxuw0r\n849R5yv4eTBATy+r9kfjvyNfvqMD1JP+G5jQeRqclZ7av7iNgANNUGc50m4TthRe\n7vC0Q/ghAgMBAAECggEAAnOWY9lY+yLhp1QYhksrSnWBv/rK4pZ1C7g6ztfhudHf\nmHzT7z7UocEYs4AT4TYdIh1Epa/qBOh8RNc1bhp+rAFAzEfMUS2+U7zGXCaAlgvi\ndYLR+nDl3QXkWW/kkU9FDRijXhJ9K4tLO9sMjsMSW4YlXRi+gb8sH6AOo5/L6DjR\n7ZM01ftzTN967+mrdnvHbgZt69iCoygYpyYiAb076EUdPPgAd8Be6ogohKhvMWA6\ns/tfXtiHpqzBieB2EUnsCS5QwfxReG6cNPT9BWXkq4MAKGy7Ctcz3LYpqcC5aDdb\n1NGp3Q8vaHcZ30zJccLUTWPoHs2COkDtSW18p+GMAQKBgQDY1gaVwYXNS7GG+rdM\n616qCuVyDbbQwArTuK3I4jhnnjB9hj9uMbIAInXhPJWTKRXybMbfU2mCsI5VEEJW\ndibfZpQ0J7Nn19OCMdLG4JTB59sZnN/qHLLVkkiy2+zxquciriEwaG2T5Q0DTrLG\n3CbgkSzOIdU0tObizszaE7RfoQKBgQDJ2DQhP3r8v5/qa8Yf4vCXzIVywqUb31kr\nfqOaDzp19Cd4e19gl/LK2tAEjMiA+uKmsMhmlIA1vfhcSrR/fzNkuKCsI9A/yhl+\nJ8WqtmblFdCklfsHSzrPcE5/P4rSZMXuB3Ma6ZBnjYLuYzaf9LIhG9XSPkvOlMXr\nqR7dY2rIgQKBgHU17EVTYOKCgio2qJL5wCgmz7SBWUsqJDAiaj5mmprYVdnkkbEd\nR9zuw83HFAuCcAylZDMgQa6VhbrRmSpnn8evCXnP5BjD/98m04sRpxfSHwuPUzKX\n5Mux0X5th31zJpIGkoY6TNFfRVN+XQFFy/YkQ5YBj+B30T7VDsLrV9tBAoGBAJTP\nNhexEn5W1JJipK4LhS+VFGm4QTwcXURo2DsTsRkXSSZVZsrzG4gc7DH+jTAyR3l3\najfekeuNRBbe6NX6tKw0RhjDSpxM5qCQt/WVBqUsgSdmf60v9IrNFMJR2Yoly5si\nmOUlf1YpCXexY6toHw+z0t9vGDqUipqkk+HKkwaBAoGAS1l8SQAibo4riKdbuwMO\nafyeagq/nvwmGeD6DcjugLI72aaSSbxgFVgBTxMN6cEdsAcS5wtTnco3+qUYRbSM\n4UAbAK+lb3xEtCXf0r0QpLZa77WuHU3Kyfb+2dmgO03psDMOdLXBJVDjaEkVDLIf\ntUe+YQaOZ7axnBd+A6GpOVc=\n-----END PRIVATE KEY-----\n"
  const GOOGLE_SHEET_ID = "1vMsgHt84RJoHLV67l6ca1jvoigyPErPfGkGVqXVkCf0"

  const spreadsheetId = GOOGLE_SHEET_ID;
  const auth = new JWT({
    email: GOOGLE_SERVICE_ACCOUNT_EMAIL,
    key: GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'), // Replace any escaped newline characters
    scopes: "https://www.googleapis.com/auth/spreadsheets",
  });
  const googleSheets = google.sheets({ version: "v4", auth });

  let queryString = '';
  const keys = Object.keys(req.query);
  keys.forEach((key, index) => {
    if (index > 0) queryString += '&';  // Use '?' as delimiter
    queryString += `${key}=${encodeURIComponent(req.query[key])}`;
  });

  const currentDate = new Date();
  currentDate.setMinutes(currentDate.getMinutes() - 3);
  const currentFormattedTime = currentDate.toISOString().slice(0, 10);
  const data = {
    values: [
      [
        currentFormattedTime,
        req.query["WF_Name"],
        req.query["WF_No"],
        req.query["WF_ClinicName"],
        req.query["WF_Birthdate"],
        req.query["WF_address"],
        'クーポンコードなし',
        req.query["WF_email"],
        req.query["WF_phone_number"],
      ]
    ]
  };
  try {
    const getRows = await googleSheets.spreadsheets.values.get({
      auth,
      spreadsheetId,
      range: 'シート1', // Replace with your actual sheet name
    });

    const rows = getRows.data.values || [];

    // Search for existing data with the same currentFormattedTime, WF_Name, and WF_No
    let rowIndex = -1;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] === currentFormattedTime && (rows[i][1] === req.query["WF_Name"] || rows[i][4] === req.query["WF_Birthdate"] || rows[i][5] === req.query["WF_address"] || rows[i][7] === req.query["WF_email"] || rows[i][8] === req.query["WF_phone_number"])) {
        rowIndex = i;
        break;
      }
    }

    console.log("data1", data);
    const redirect_url = 'https://demo.services.docusign.net/webforms-ux/v1.0/forms/c4db2471f86bc8475ea9734fd7b2f7ad#'+queryString;

    if (rowIndex !== -1) {
      // Update existing data
      await googleSheets.spreadsheets.values.update({
        spreadsheetId: spreadsheetId,
        range: `シート1!A${rowIndex + 1}`, // A1 notation of the row to be updated
        valueInputOption: 'RAW',
        resource: data
      });
      console.log('Data updated successfully:');
      res.redirect(redirect_url);
    } else {
      // Append new data
      await googleSheets.spreadsheets.values.append({
        spreadsheetId: spreadsheetId,
        range: "シート1",
        valueInputOption: 'RAW',
        resource: data
      });
      console.log('Data appended successfully:');
      res.redirect(redirect_url);
    }
  } catch (err) {
    console.error('The API returned an error:', err);
    res.status(500).send("Error appending or updating data");
  }
});

app.listen(5000, () => {
  console.log('Server started on http://localhost:5000'); // Listen on localhost:3000
});

module.exports = app; // Export the app for potential testing or further integration
