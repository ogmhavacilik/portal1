const { groupMultiLocationRowsHelper } = require('./test-group-helper.cjs');

const rawParsedRows = [
  [ "1", "Boyler", "200 litre", "-", "1", "Kazan Dairesi", "FAAL", "EVET" ],
  [ "2", "Yangın Söndürme Cihazı", "Köpüklü 25 kg Tekerli", "-", "2", "Uçuş Hattı", "FAAL", "EVET" ],
  [ "", "", "", "-", "1", "Uçuş Hattı", "FAAL", "EVET" ]
];

const result = groupMultiLocationRowsHelper(rawParsedRows, 1, 5, 4, 0, 2, 3);
console.log(result);
