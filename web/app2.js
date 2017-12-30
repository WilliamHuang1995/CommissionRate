var rABS = true; // true: readAsBinaryString ; false: readAsArrayBuffer
function handleFile(e) {
  var files = e.target.files, f = files[0];
  var reader = new FileReader();
  reader.onload = function(e) {
    var data = e.target.result;
    if(!rABS) data = new Uint8Array(data);
    var workbook = XLSX.read(data, {type: rABS ? 'binary' : 'array'});

    /* DO SOMETHING WITH workbook HERE */
    var sheet = workbook.SheetNames[0];
    var o = XLSX.utils.sheet_to_csv(workbook.Sheets[sheet])
  };
  if(rABS) reader.readAsBinaryString(f); else reader.readAsArrayBuffer(f);
}
var input_dom_element = document.getElementById('input')
input_dom_element.addEventListener('change', handleFile, false);
