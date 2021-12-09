/* Macro for counting nuclei and Ki67 on RGB IHC image  - SK/Janina Janssen P44 2019-04 VTB
 * Settings for 3DHisterch Pathoscanner 20x/20x
 * exported with 3DHistech SlideConverter 1:1 tif, uncompressed jpg --> 0.496µm/Px
 * and converted to png with FIJI with 80% scaling factor --> 0.62µm/Px 
 *
 * SK / VetImaging-VetBiobank / VetCore / Vetmeduni Vienna 2019
 * This research was supported using resources of the VetCore Facility (VetImaging | VetBioBank) of the University of Veterinary Medicine Vienna
 */

setForegroundColor(0, 0, 0);
setBackgroundColor(255, 255, 255);
scale_um = 0.62;

//clear the log
print("\\Clear");

// clear previous results
run("Clear Results");

// set measurement parameters
run("Set Measurements...", "area shape limit display redirect=None decimal=1");

var dir=getDirectory("Choose a Directory");
var plain_dir_name=File.getName(dir);

//get data list and count number of images
var list = getFileList(dir);
var img_list_length=0;
for (filenumber=0; filenumber<list.length; filenumber++) {

     if (endsWith(list[filenumber], ".png")){
     		img_list_length=img_list_length+1;
     }}

// create Folder for Nuclei, Ki67 and cropped images
var nucleiDir=dir + "\Nuclei\\";
File.makeDirectory(nucleiDir);
var ki67Dir=dir + "\Ki67\\";
File.makeDirectory(ki67Dir);
var croppedDir=dir + "\cropped_image\\";
File.makeDirectory(croppedDir);


//define array-variables
var plain_title=newArray(img_list_length);
var total_area=newArray(img_list_length);

var nuclei_count=newArray(img_list_length);
var nuclei_area=newArray(img_list_length);
var nuclei_avsize=newArray(img_list_length);
var nuclei_circularity=newArray(img_list_length);
var nuclei_solidity=newArray(img_list_length);

var Ki67_count=newArray(img_list_length);
var Ki67_area=newArray(img_list_length);
var Ki67_avsize=newArray(img_list_length);
var Ki67_circularity=newArray(img_list_length);
var Ki67_solidity=newArray(img_list_length);

var Ki67_ratio=newArray(img_list_length);
var jpg_number=0;


//take one after another image and process
for (filenumber=0; filenumber<(list.length); filenumber++) {
     if (endsWith(list[(filenumber)], ".png")){
             run("Collect Garbage");
             open(dir+list[(filenumber)]); 


//get title of image and rename for results list
var imageTitle = getTitle();
plain_title[jpg_number] = getTitle();
plain_title[jpg_number] = replace(plain_title[jpg_number], "\\.png", "");


//set scale for image
run("Set Scale...", "distance=1 known=" + scale_um + " unit=µm global");

// Set area manually to analyze
do {
	setTool("freehand");
	waitForUser("Outer Boundaries", "Draw boundary of tumor, then click OK.");
} while(selectionType() == -1);

getStatistics(total_area_temp, mean_l, min_l, max_l, std_l, histogram_l);
total_area[jpg_number]=total_area_temp;

run("Clear Outside");
run("Crop");
saveAs(".jpg", croppedDir + plain_title[jpg_number] + "-cropped.jpg");

//duplicate original image for H-Overlay
run("Duplicate...", " ");
H_Overlay = getImageID();
selectImage(H_Overlay);
rename(plain_title[jpg_number] + "-H_Overlay");

selectWindow(imageTitle);

// colour deconvolution
run("Colour Deconvolution", "vectors=[H DAB]");

// rename images
//close green channel
selectWindow(imageTitle + "-(Colour_3)");
run("Close");

selectWindow(imageTitle + "-(Colour_1)");
rename(plain_title[jpg_number] + "-Nuclei");
run("8-bit");

selectWindow(imageTitle + "-(Colour_2)");
rename(plain_title[jpg_number] + "-Ki67");
run("8-bit");



//** count nuclei 
//duplicate image-nuclei and do analysis on copy
selectWindow(plain_title[jpg_number] + "-Nuclei");
run("Duplicate...", " ");
copy = getImageID();
selectImage(copy);
rename(plain_title[jpg_number] + "-copy");

selectWindow(plain_title[jpg_number] + "-copy");

// background subtraction, enhance contrast
run("Subtract Background...", "rolling=50 light");
run("Enhance Contrast...", "saturated=0.3 normalize");

// run("Bilateral Filter", "spatial=10 range=50");
selectWindow(plain_title[jpg_number] + "-copy");
rename(plain_title[jpg_number] + "-Nuclei_count");

// set threshold and create selection
setThreshold(0, 190);
run("Threshold...");

//create selection and draw boundaries to image
run("Convert to Mask");
run("Watershed");

// analyze particles, limited by size, add to roi manager
run("Analyze Particles...", "size=13-130 display clear include summarize add in_situ");

//save detailed results of nuclei
selectWindow("Results");
saveAs("Text", nucleiDir + plain_title[jpg_number] + "_nuclei details.xls");

//save results in arrays
selectWindow("Summary");
IJ.renameResults("Results");
      nuclei_count[jpg_number] = getResult("Count",0);
      nuclei_area[jpg_number]=getResult("Total Area",0);
      nuclei_avsize[jpg_number]=getResult("Average Size",0);
      nuclei_circularity[jpg_number]=getResult("Circ.",0);
      nuclei_solidity[jpg_number]=getResult("Solidity",0);
 
// save overlay image
selectWindow(plain_title[jpg_number] + "-H_Overlay");
roiManager("Show None");
roiManager("Show All without labels");
roiManager("Set Color", "00FFFF");
roiManager("Set Line Width", 2);
run("Flatten");
saveAs(".jpg", nucleiDir + plain_title[jpg_number] + "-H_Overlay.jpg");
roiManager("Delete");
run("Close");


// close unused images
if (isOpen(plain_title[jpg_number] + "-Nuclei")) { selectWindow(plain_title[jpg_number] + "-Nuclei"); run("Close"); }
if (isOpen(plain_title[jpg_number] + "-copy")) { selectWindow(plain_title[jpg_number] + "-copy"); run("Close"); }
if (isOpen(plain_title[jpg_number] + "-Nuclei_count")) { selectWindow(plain_title[jpg_number] + "-Nuclei_count"); run("Close"); }
if (isOpen(plain_title[jpg_number] + "-H_Overlay")) { selectWindow(plain_title[jpg_number] + "-H_Overlay"); run("Close"); }
if (isOpen("Results")) { selectWindow("Results"); run("Close"); }


//** count DAB positive cells
//duplicate original image for Overlay
selectWindow(imageTitle);
run("Duplicate...", " ");
Ki67_Overlay = getImageID();
selectImage(Ki67_Overlay);
rename(plain_title[jpg_number] + "-Ki67_Overlay");

//take color deconvoluted Ki67-channel for analysis
selectWindow(plain_title[jpg_number] + "-Ki67");

// background subtraction, enhance contrast
run("Subtract Background...", "rolling=50 light");
run("Enhance Contrast...", "saturated=0.1 normalize");

rename(plain_title[jpg_number] + "-Ki67_count");

// set threshold and create selection
setThreshold(0, 160);
run("Threshold...");

//create selection and draw boundaries to image
run("Convert to Mask");
run("Watershed");

// analyze particles, limited by size, add to roi manager
run("Analyze Particles...", "size=13-130 display clear include summarize add in_situ");

//save detailed results of Ki67 positive Nuclei
selectWindow("Results");
saveAs("Text", ki67Dir + plain_title[jpg_number] + "_Ki67 details.xls");

//save results in arrays
selectWindow("Summary");
IJ.renameResults("Results");
      Ki67_count[jpg_number] = getResult("Count",0);
      Ki67_area[jpg_number]=getResult("Total Area",0);
      Ki67_avsize[jpg_number]=getResult("Average Size",0);
      Ki67_circularity[jpg_number]=getResult("Circ.",0);
      Ki67_solidity[jpg_number]=getResult("Solidity",0);
 
// save overlay image
selectWindow(plain_title[jpg_number] + "-Ki67_Overlay");
roiManager("Show None");
roiManager("Show All without labels");
roiManager("Set Color", "FF0000");
roiManager("Set Line Width", 2);
run("Flatten");
saveAs(".jpg", ki67Dir + plain_title[jpg_number] + "-Ki67_Overlay.jpg");
roiManager("Delete");
run("Close");

//calculate ratio Ki67 vs all nuclei
Ki67_ratio[jpg_number]=(100/nuclei_count[jpg_number]*Ki67_count[jpg_number]);

//close open windows
if (isOpen(plain_title[jpg_number] + "-Ki67")) { selectWindow(plain_title[jpg_number] + "-Ki67"); run("Close"); }
if (isOpen(plain_title[jpg_number] + "-Ki67_count")) { selectWindow(plain_title[jpg_number] + "-Ki67_count"); run("Close"); }
if (isOpen(plain_title[jpg_number] + "-Ki67_Overlay")) { selectWindow(plain_title[jpg_number] + "-Ki67_Overlay"); run("Close"); }
if (isOpen(imageTitle)) { selectWindow(imageTitle); run("Close"); }
if (isOpen("Results")) { selectWindow("Results"); run("Close"); }
jpg_number=jpg_number+1;

 }
}

//selectWindow("Results");
run("Measure");
run("Clear Results");

//Read out values of arrays into Results-Table and save Results-table as csv
for (result_output=0; result_output<img_list_length; result_output++) {

setResult("Image", result_output, plain_title[result_output]); 
setResult("Total Area (µm²)", result_output, total_area[result_output]); 
setResult("Nuclei count", result_output, nuclei_count[result_output]); 
setResult("Ki67 count", result_output, Ki67_count[result_output]); 
setResult("Ki67 ratio (%)", result_output, Ki67_ratio[result_output]); 
setResult("Nuclei average size (µm²)", result_output, nuclei_avsize[result_output]); 
setResult("Ki67 average size (µm²)", result_output, Ki67_avsize[result_output]); 

/*
	further possible variables, not used here: 
	nuclei_area
	nuclei_circularity
	nuclei_solidity
	Ki67_area
	Ki67_circularity
	Ki67_solidity
*/
}

selectWindow("Results");
saveAs("Text", dir + "Summary-"+plain_dir_name+".xls");

// close all windows
run("Close All");
if (isOpen("Results")) { selectWindow("Results"); run("Close"); } 
if (isOpen("Threshold")) { selectWindow("Threshold"); run("Close"); } 
if (isOpen("ROI Manager")) { selectWindow("ROI Manager"); run("Close"); }    

// indicate end of macro processing in the log file by a text   
print("Macro is finished");
