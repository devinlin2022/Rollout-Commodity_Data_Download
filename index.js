function redirectFunction() {
  //location.replace("https://docs.google.com/spreadsheets/d/1q6tLUMal9t1EafNgs5ZeI-623e7Wh-bS-Sv1RwRIPQg/edit#gid=71191126")
    const dt = document.getElementById('01').value;
    const sku = document.getElementById('02').value;
    const qty = document.getElementById('03').value;
    const dict_values = {dt,sku,qty}
    const s = JSON.stringify(dict_values)
    console.log(s)

    $.ajax({
      data : s,
      type : 'POST',
      url : '/test',
      contentType:"application/json",
      success: function(response){
        console.log('success');},
      error: function (response){
            console.log('error');}
    });

}

