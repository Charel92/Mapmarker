function startMapService() {
  function_block:{
    var doc = DocumentApp.getActiveDocument();
    var tables = doc.getBody().getTables();
    var destinationMap = Maps.newStaticMap();
    try{
      var destinations=searchTable("Station",tables); 
    }
    catch(e){
      DocumentApp.getUi().alert("You must write Station in the table with your locations");
      break function_block;
    }
    for (var i = 1; i < destinations.getNumRows(); i++) {
      destinationMap.setMarkerStyle(Maps.StaticMap.MarkerSize.MID,Maps.StaticMap.Color.RED,i);
      var x =destinations.getCell(i,0).editAsText().getText();
      if(destinations.getCell(i,0).editAsText().getText()!=="" && destinations.getCell(i,0).editAsText().getText()!==" "){
        destinationMap.addMarker(destinations.getCell(i,0).editAsText().getText());
      }
     
      // This is for debugging
      //var text=destinations.getCell(i,0).editAsText().getText()
    }
    var body = doc.getBody();
    // Retrieve an image from the MapURL.
    var resp = UrlFetchApp.fetch(destinationMap.getMapUrl());
    var image = resp.getBlob();
    var body = doc.getBody();
    try{
      var oImg = body.findText("<<MAP>>").getElement().getParent().asParagraph();
    }
    catch(e){
      DocumentApp.getUi().alert("You must write <<MAP>> on the place where you want your Map.");
      break function_block;
    }
    
    oImg.clear();
    oImg = oImg.appendInlineImage(image);
    // This is for debugging
    // DocumentApp.getUi() // Or DocumentApp or FormApp.
    // .alert(text+'You clicked the first menu item!'+destinationMap.getMapUrl());
  }  
}

function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Add Map', 'startMapService')
      .addToUi();
}

function onInstall(e) {
  onOpen(e);
}
function searchTable(name, tables){
  //var length = tables.length;
  for (var i=0; i < tables.length ; i++ ){
    if(tables[i].getCell(0,0).editAsText().getText()===name){
      return tables[i];
    }
  }
  
  throw'myException';
  
  //catch(e){
  //  DocumentApp.getUi().alert("You must write Station in the table with your locations");
  //}
  
  

}



