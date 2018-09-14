// when the user presses submit
$('#submit').click(function() {
  // check if file is uploaded
  var selectedFile = document.getElementById('input').files[0];
  // if file is selected
  if (selectedFile) {
    //use workbook
    var reader = new FileReader();
    reader.onload = function(e){
      var data = document.getElementById('input').result;
      if(!rABS) data = new Uint8Array(data);
      var workbook = XLSX.read(data, {type: rABS ? 'binary' : 'array'});
    };
    if(rABS) reader.readAsBinaryString(selectedFile); else reader.readAsArrayBuffer(selectedFile);
  } else {
    alert("No file selected!")
  }

})


var rABS = true; // true: readAsBinaryString ; false: readAsArrayBuffer
