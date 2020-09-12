#include "./csv.jsx"

// TEMPLATE FILE PATH
var myTemplate = File((new File($.fileName)).parent + "/templates/Quarterly Report - Wireframe 4.0.indd");
// OPEN TEMPLATE
var newDoc = app.open(myTemplate);
// ACTIVATE
var myDoc = app.activeDocument;
// GET STORIES
var myStories = myDoc.stories;
// GET TEXT FRAMES
var myTextFrames = myDoc.textFrames;
// GET FRAMES
var myRectangleFrames = myDoc.rectangles;


// FUNCTION FOR READING IN EXCEL FILE
function readCSVData() {
        //var f = File.openDialog ("Please select the CSV File…", "*.txt", false),
        var f = File((new File($.fileName)).parent + "/data/Q2 2020 QR Training.csv");
        
        // if no file, return
        if ( !f ) return;
        // open the file and read
        f.open('r');

        // READ IN CSV
        var textLines = CSV.reader.read_in_txt();
       
       var csvData = -1;
        if(textLines !== null){
            var csvData = CSV.reader.textlines_to_dict(textLines,",");           
        }

        return csvData;
};

function activeLabels(){
        $.write(myTextFrames.length);
        var n = 0;
        for (var i = 0; i < myTextFrames.length; i++) {
                if (myTextFrames[i].label != ""){
                        $.write(myTextFrames[i].label + "\n");
                        n += 1;
                        $.write(String(n) + "\n")
                }
              
        }
}

function countMasterFrames(){
        var masterSpreads = app.activeDocument.masterSpreads;
        totalMasterFrames = 0;
        for (var i = 0; i < masterSpreads.length; i++){
                totalMasterFrames += masterSpreads.item(i).textFrames.length;
        }
        return totalMasterFrames;
}


function getFileName(fullPath){
        var filename = fullPath.name.replace(/^.*[\\\/]/, '');
        return filename;
}


function getPhotoNames(propertyName){
        // GET IMAGE FOLDER FOR PROPERTY
        // TODO: THROW ERROR WHEN PROPERTY NOT FOUND
        var myFolder = Folder((new File($.fileName)).parent + "/img/" + propertyName);
        var images = myFolder.getFiles();
        var pnamePathMap = {}        
        for (var i = 0; i < images.length; i++) {        
                pnamePathMap[getFileName(images[i])] = images[i];
        }
        return pnamePathMap;
}

function getStaticIndex(position, label, myTextFrames, textFramesPerPage){
        for(var i = 0; i < textFramesPerPage; i++){
               if(myTextFrames[position+i].label == label){
                        return position+i;
               }
        }
        //TODO: THROW ERROR
}

function getPropertyImagePath(photoMap, img_name){
        var pm_size = photoMap.length;
        for (var i = 0; i < pm_size; i++){
            if (photoMap[i][img_name])
                   return photoMap[i][img_name];  
            }    
    }


function buildQuarterlyReport( reportData ){
        
        // COUNT MASTER FRAMES 
        var masterFramesCount = countMasterFrames();
        var textFramesPerPage = myTextFrames.length - masterFramesCount;
       
       var report_length = reportData.length;       
        // CREAT ALL PAGES
        
        for (var i = 0; i < reportData.length - 1; i++){
                //app.layoutWindows[0].activePage.duplicate (LocationOptions.AFTER, app.layoutWindows[0].activePage);
                app.layoutWindows[0].activeSpread.duplicate (LocationOptions.AFTER, app.layoutWindows[0].activeSpread);
        }

//~         for (var i = 0; i < myTextFrames.length; i++){
//~                 $.write(String(i)+"  "+myTextFrames[i].label + "\n");
//~         }
        
        
        var photoMap = []
        for (var j = 0; j < reportData.length; j++) {
               var property = reportData[j];
               var property_name = property.Property;
               var prop_photoMap = getPhotoNames(property_name); 
               photoMap.push(prop_photoMap);      
            }
        
        
//~         $.write(myTextFrames.length + "\n");
//~         $.write(reportData.length + "\n");
        for (var j = 0; j < reportData.length; j++) {
               
               var property = reportData[j];
               var property_name = property.Property;
               
               for ( var k = 0; k <  textFramesPerPage; k++) {
                       var i = j*textFramesPerPage+k;
                       var textframe = document.textFrames.item(i);
                       //$.write(myTextFrames[i].label + "\t" + String(i) + "\n");
                        switch (myTextFrames[i].label) {  
                                case "ASSET_NAME":
                                        textframe.parentStory.contents = property["Property"];
                                        break;
                                        
                                case "ASSET_TYPE":
                                        myTextFrames[i].contents = property["Asset Class"];
                                        break;
                                        
                                case "MSA":
                                        var citystate = property["City"] + ", " + property["State"] + " | ";
                                        var msa = property["MSA"] + " MSA";
                                        myTextFrames[i].contents = citystate + msa;
                                        break;
                                
                                case "ACQUISITION":
                                        myTextFrames[i].contents = property["Acq. Date"];
                                        break;
                                
                                case "PRICE_SF":
                                        if (property["Asset Class"].toLowerCase().localeCompare("multi-family") != 0){
                                                myTextFrames[i].contents = property["Buildings"];
                                                var index = getStaticIndex(j*textFramesPerPage, "STATIC_PRICESF", myTextFrames, textFramesPerPage)
                                                myTextFrames[index].parentStory.contents = "Buildings:";
                                        }else {
                                                myTextFrames[i].contents = property["Price Per SF"] + "/sf";
                                        }
                                        break;
                                
                                case "PRICE_UNIT":
                                        if (property["Asset Class"].toLowerCase().localeCompare("multi-family") != 0){
                                                myTextFrames[i].parentStory.contents = String(property["Price Per SF"] + "/sf");
                                                var index = getStaticIndex(j*textFramesPerPage, "STATIC_PRICEUNIT", myTextFrames, textFramesPerPage)
                                                myTextFrames[index].parentStory.contents = "Price/sf:";
                                        }else {
                                                myTextFrames[i].contents = property["Price Per Unit"];
                                        }
                                        break;
                                
                                case "PURCHASE_PRICE":
                                        myTextFrames[i].contents = property["Purchase Price"];
                                        break;
                                
                                case "YEAR_BUILT":
                                        myTextFrames[i].contents = property["Year Built"];
                                        break;
                                        
                                case "OCCUPANCY_RATE":
                                        myTextFrames[i].contents = property["Quarter End Occupancy"];
                                        break;
                                        
                                case "PAY_RATE":
                                        myTextFrames[i].contents = property["Annualized Distribution"];
                                        break;
                                
                                case "UNITS":
                                         if (property["Asset Class"].toLowerCase().localeCompare("multi-family") != 0){
                                                myTextFrames[i].contents = property["Square Feet"] + " sf";
                                                var index = getStaticIndex(j*textFramesPerPage, "STATIC_UNITS", myTextFrames, textFramesPerPage);
                                                myTextFrames[index].parentStory.contents = "SF:";
                                        }else {
                                                myTextFrames[i].contents = property["Units"];
                                        }
                                        break;
                               
                                case "PROPERTY_SUMMARY":
                                        myTextFrames[i].parentStory.contents =  property["Property Summary"];
                                        break;
                                
                                case "OPERATIONS_SUMMARY_1":
                                        if (property["Operations-1"] == 0) break;
                                        myTextFrames[i].parentStory.contents =  property["Operations-1"];
                                        if (property["Operations-2"] == 0) break;
                                        myTextFrames[i].parentStory.insertionPoints[-1].contents=  "\r" +property["Operations-2"];
                                        if (property["Operations-3"] == 0) break;
                                        myTextFrames[i].parentStory.insertionPoints[-1].contents=  "\r" + property["Operations-3"];
                                         if (property["Operations-4"] == 0) break;
                                        myTextFrames[i].parentStory.insertionPoints[-1].contents=  "\r" + property["Operations-4"];
                                         if (property["Operations-5"] == 0) break;
                                        myTextFrames[i].parentStory.insertionPoints[-1].contents=  "\r" + property["Operations-5"];
                                         break;
                                        
                                default:                                        
                                        break;
                  
                        }
                } 
        }
                
        // INSERT IMAGES
        for(var i = 0; i < myDoc.rectangles.length; i++){
                $.write(myDoc.rectangles[i].label + "\n");
                var thisRect = myDoc.rectangles[i];
                var prop_num = Math.floor(i / 4);
                switch (thisRect.label) {
                          case "MAIN_IMAGE":
                                        var path = photoMap[prop_num]["main_banner.jpg"];
                                        thisRect.place(File(path));
                                        thisRect.fit (FitOptions.CONTENT_AWARE_FIT);
//~                                         thisRect.fit (FitOptions.PROPORTIONALLY);
//~                                         thisRect.fit (FitOptions.CENTER_CONTENT);
                                        break;
                               case "LEFTBAR_UPPER_IMAGE":
                                        var path = photoMap[prop_num]["side_banner_mid.jpg"];
                                        thisRect.place(File(path));
                                        thisRect.fit (FitOptions.CONTENT_AWARE_FIT);
//~                                         thisRect.fit (FitOptions.PROPORTIONALLY);
//~                                         thisRect.fit (FitOptions.CENTER_CONTENT);
                                        
                                        break;
                               case "LEFTBAR_LOWER_IMAGE":
                                        var path = photoMap[prop_num]["side_banner_bottom.jpg"];
                                        thisRect.place(File(path));
                                        thisRect.fit (FitOptions.CONTENT_AWARE_FIT);
//~                                         thisRect.fit (FitOptions.PROPORTIONALLY);
//~                                         thisRect.fit (FitOptions.CENTER_CONTENT);
                                        break;
                               case "UPPER_LEFT_IMAGE":
                                        var path = photoMap[prop_num]["side_banner_top.jpg"];
                                        thisRect.place(File(path));
                                        thisRect.fit (FitOptions.CONTENT_AWARE_FIT);
//~                                         thisRect.fit (FitOptions.PROPORTIONALLY);
//~                                         thisRect.fit (FitOptions.CENTER_CONTENT);
                                        break;
                                default:
                                        break;
                }
        }
        
        
}

function main(){
        
        // READ DATA
        var reportData = readCSVData();
        if (reportData == -1) $.write("ERROR");
        buildQuarterlyReport(reportData);
        
}



main();


