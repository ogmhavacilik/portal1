const https = require('https');
const fs = require('fs');

const url = "https://docs.google.com/spreadsheets/d/1_a8BrcQlerc9qLJUbITFl4qKafgZZOxU/export?format=xlsx";

https.get(url, (res) => {
  if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
    https.get(res.headers.location, (res2) => {
      const file = fs.createWriteStream("test.xlsx");
      res2.pipe(file);
      file.on('finish', () => file.close());
    });
  } else {
    const file = fs.createWriteStream("test.xlsx");
    res.pipe(file);
    file.on('finish', () => file.close());
  }
});
