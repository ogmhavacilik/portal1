async function run() {
  const url = "https://script.google.com/macros/s/AKfycby5t4cI5n7r2-N-W4wK9_F5C1B8V4K-iF5t3T-uU8pS1wM3_R-fV_7h/exec?action=listTechizatExcelsFromDrive&folderId=1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP";
  const res = await fetch(url);
  const data = await res.json();
  console.log(JSON.stringify(data.excels.map(f => f.name), null, 2));
}
run();
