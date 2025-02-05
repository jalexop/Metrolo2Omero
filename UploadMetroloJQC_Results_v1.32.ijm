// @String (visibility=MESSAGE, value="<html><h1>Upload MetroloJ QC results to Omero</h1></html>", required=false) head
// @String (visibility=MESSAGE, value="<html><h2>Omero Login Info</h2></html>", required=false) msg
// @String(label="Host", value='omero.quarep.org') omrsrv
// @Integer(label="Port", value=4064) omrport
// @String(label="Omero Username", style="Text Field") omrusr
// @String(label="Password", style='password', persist=false) omrpwd
// @String (visibility=MESSAGE, value="", required=false) msg11
// @String (visibility=MESSAGE, value="<html><h2>Upload Information</html></h2>", required=false) msg2
// @File (label="Choose a root Directory with raw images ", style="directory") dir
// @String (choices={"wg05", "wg04", "wg03", "wg01", "wg06", "---"}, style="listBox") Workgroup
// @String (choices={"LSM", "WFM", "SD"}, style="listBox") Modality
// @String(label="Experiment Number", style="0001") ZZZZ
// @Integer (label="Instrument Indentifier", value=00000000, persist=false) InstrumentIdent
// @String (choices={"PSF", "Illumination", "CoRegistration", "---", "---"}, style="listBox") QCType
// @String (choices={"FWHM", "---", "---", "---", "---"}, style="listBox") QCSubType




/*omero_macro-extensions-1.4.0.jar

 * Removed Dialogue Components
 * // @String (choices={"---", "100x", "63x", "60x", "40x"}, style="listBox") Magnification
 * // @Double (value=1.4, min=0.4, max=1.5, stepSize=0.01, persist=false, style="slider,format:0.00") NA
 * 
 * // @Boolean(label="Batch Mode", value=false) batch
 * 
 * // @java.util.Date AcquisitionDate
 * // @String(label="Omero Project", style="Text Field") omrProject
 */

////////////////////////////////////////////////////////////////////////////////////////////////
/*
 * 				Macro written by Dr. I. Alexopoulos
 * The author of the macro reserve the copyrights of the original macro.
 * However, you are welcome to distribute, modify and use the program under 
 * the terms of the GNU General Public License as stated here: 
 * (http://www.gnu.org/licenses/gpl.txt) as long as you attribute proper 
 * acknowledgement to the author as mentioned above.
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 ***********************************************************************************************
 * Description of macro
 * --------------------
 * This macro is used for uploading on an omero server the results of the PSF 
 * analysis with MetroloJ QC. It uploads the raw files analysed by MetroloJ with rois 
 * of the analysed beads. It also generates a table (attached at the dataset level) with
 * the FWHM measurements as well as the R2 for each analysed bead.
 * It also uploads (attaches) the results as a tab-separated .txt file and the summary.pdf file 
 * at the Dataset level.
 * 
 * 			 Dependencies
 * Tested with the latest versions of plugings as mentioned in:
 * https://github.com/GReD-Clermont/simple-omero-client
 * (Currently tested with version 5.19.0: simple-omero-client-5.19.0-jar-with-dependencies.jar)
 * and
 * https://github.com/GReD-Clermont/omero_macro-extensions
 * (currently tested with omero_macro-extensions-1.4.0.jar)
 * 
 * as well as the latest version of Omero.insight plugin:
 * https://www.openmicroscopy.org/omero/downloads/
 * (currently tested with omero_ij-5.8.6-all.jar)
 * 
 * This plugin requires MetroloJ QC 1.3.1.1
 * https://github.com/MontpellierRessourcesImagerie/MetroloJ_QC
 * or from Fiji Update sites
 * https://sites.imagej.net/MetroloJ_QC/
 * 
 * 
 */
requires("1.54f");
Macro_version=1.32;
//NumAp=NA;
//if(Magnification=="---"){
//	Magnification="";
//}
if(Workgroup=="---"){
	Workgroup="";
}
omrProject=Workgroup+"-"+ZZZZ+"_"+Modality+"_"+QCType+"_"+QCSubType;
//Fixing the date format
//date=split(AcquisitionDate, " ");
//month=(indexOf("JanFebMarAprMayJunJulAugSepOctNovDec", date[1]))/3;
//month++;
//if(month<10){month="0"+month;}
//AcqDate=date[5]+"-"+month+"-"+date[2];
//
//User: First letter of first name and three first letters of surname (taken by Omero username)
user=substring(omrusr, 0, 4);
setBatchMode(true);

//Global variables
sep=File.separator;
lineseparator = "\n";
cellseparator = "\t";
dir=dir+sep;
dir_proc=dir +"Processed/";

//Find the local files (paths, names) by running the dedicated functions
//Parsing the results (summary.xls) into an array
countRaw=countRawFiles(dir);
RawFilePaths=FindRawFilePaths(dir, countRaw);
Raw_Names=FindRawFileNames(dir, countRaw);
ExperimentName=getExperimentName(dir_proc);
datasetName=replace(ExperimentName, "/", "");
summaryLines=split(File.openAsString(dir_proc+ExperimentName+datasetName+"_summary.xls"), lineseparator);

//  Connect to Omero server, and create a new Project (based on user's input). If the Project already exists for this user
//then the same Project is used for import.
//  A new Dataset is also created (based on the MetroloJ experiment name (folder name under the Processed Folder). If the Dataset
//within the Project, already exists, then it is the one that will be used for the import. 
// ---> Check what happens with the Omero table if the dataset exists
run("OMERO Extensions");
succesfulConnectionOmero=Ext.connectToOMERO(omrsrv, omrport, omrusr, omrpwd);
if (succesfulConnectionOmero){
	print("\\Clear");
	print ("################### Connection to Omero and Project / Dataset Management ################");
	print ("Succesfully Connected to Omero Server "+omrsrv+" on port "+omrport+" for user "+omrusr+"");
}else{
	exit("Cannot connect to Omero Server "+omrsrv+" on port "+omrport+" for user "+omrusr+"");
}
ListForUser=Ext.listForUser(omrusr);
oldProject=Ext.list("Project", omrProject);
oldTagWorkgroup=Ext.list("tags", Workgroup);
oldTagModality=Ext.list("tags", Modality);
if(oldProject!=""){
	ProjectID=oldProject;
	print("Project with ID "+ProjectID+" and name "+Ext.getName("Project", ProjectID)+" already exists and will be used.");
	ProjectTags=split(Ext.list("tag", "project", ProjectID), ",");
	TagFound=false;
	if(oldTagWorkgroup!=""){
		for(w=0;w<ProjectTags.length;w++){
			if(ProjectTags[w]==oldTagWorkgroup){
				TagFound=true;
			}
		}
		if(!TagFound){
			Ext.link("project", ProjectID, "tag", oldTagWorkgroup);
		}
		TagFound=false;
	}else{
		newTagWorkgroup = Ext.createTag(Workgroup, "Workgroup ID");
		Ext.link("project", ProjectID, "tag", newTagWorkgroup);
	}
	if(oldTagModality!=""){
		for(w=0;w<ProjectTags.length;w++){
			if(ProjectTags[w]==oldTagModality){
				TagFound=true;
			}
		}
		if(!TagFound){
			Ext.link("project", ProjectID, "tag", oldTagModality);
		}
		TagFound=false;
	}else{
		newTagModality = Ext.createTag(Modality, "Type of microscope");
		Ext.link("project", ProjectID, "tag", newTagModality);
	}
}else{
	ProjectID=Ext.createProject(omrProject, "Project for Uploading MetroloJ QC PSF Results");
	print("New project with ID "+ProjectID+" and name "+Ext.getName("Project", ProjectID)+" created.");
	if(oldTagWorkgroup!=""){
		Ext.link("project", ProjectID, "tag", oldTagWorkgroup);
	}else{
		newTagWorkgroup = Ext.createTag(Workgroup, "Workgroup ID");
		Ext.link("project", ProjectID, "tag", newTagWorkgroup);
	}
	if(oldTagModality!=""){
		Ext.link("project", ProjectID, "tag", oldTagModality);
	}else{
		newTagModality = Ext.createTag(Modality, "Type of microscope");
		Ext.link("project", ProjectID, "tag", newTagModality);
	}
}
ExistingDatasets=split(Ext.list("Dataset", "Project", ProjectID), ",");
foundDataset=false;
if(ExistingDatasets.length == 0){
	DatasetID=Ext.createDataset(datasetName, "", ProjectID);
	print("Dataset with ID "+DatasetID+ " and name "+Ext.getName("Dataset", DatasetID)+" created under project "+Ext.getName("Project", ProjectID)+" .");
}else{
	for(w=0;w<ExistingDatasets.length;w++){
		checkedDatasetName=Ext.getName("Dataset", ExistingDatasets[w]);
		if(checkedDatasetName== datasetName){
			DatasetID=ExistingDatasets[w];
			foundDataset=true;
		}
	}
	if(foundDataset){
		print("Dataset with ID "+DatasetID+ " and name "+Ext.getName("Dataset", DatasetID)+" already exists under project "+Ext.getName("Project", ProjectID)+" and will be used.");
	}else{
		DatasetID=Ext.createDataset(datasetName, "", ProjectID);
		print("Dataset with ID "+DatasetID+ " and name "+Ext.getName("Dataset", DatasetID)+" created under project "+Ext.getName("Project", ProjectID)+" .");
	}
}
//Start with the Raw files: Find beads, coordinates and reformat the results of MetroloJ for upload to Omero
print("\n\nStart with raw files....\n");
run("Clear Results");
NumAp=getObjectiveNA(summaryLines);
Pinhole=getPinhole(summaryLines);
AnalysedBeads=false;
for (r=0; r<RawFilePaths.length; r++){
	//open(RawFilePaths[r]);
	run("Bio-Formats Importer", "open=["+RawFilePaths[r]+"] autoscale color_mode=Default view=Hyperstack stack_order=XYCZT");
	img=getImageID();
	roiManager("reset");
	run("Clear Results");
	selectImage(img);
	rename(Raw_Names[r]);
	//The raw Image will be uploaded to Omero
	ImageID=Ext.importImage(DatasetID);
	print("Import to Omero for image ID " + ImageID + ": " +Raw_Names[r]);

	searchString=Raw_Names[r];					
	RawNameSearch=Raw_Names[r];
	searchString=replace(searchString, " - ", "-");  //MetroloJQC Processes the Image Names and removes spaces when the names are combined with the beadID at the summary.xls
	noBeads=NoOfBeads(summaryLines, RawNameSearch);
	if(noBeads>0){
		AnalysedBeads=true;
	
		resultsCounter=0;
		X_l=newArray(4);
		Y_l=newArray(4);
		Z_l=newArray(4);
		for (b=0;b<noBeads;b++){
			linesFound = getBeadMeasurements(summaryLines, searchString, b);
			X_l=split(linesFound[0], cellseparator);
			Y_l=split(linesFound[1], cellseparator);
			Z_l=split(linesFound[2], cellseparator);
			
			run("Select None");
			makeRoi(parseInt(X_l[3]), parseInt(Y_l[3]), parseInt(Z_l[3]), img);
			Roi.setPosition(0, parseInt(Z_l[3]), 0);
			roiManager("add");
			//Rename the bead Rois with the beads' names
			CurrentRoi=roiManager("count")-1;
			roiManager("select", CurrentRoi);
			roiManager("rename", Raw_Names[r]+"_bead"+b);
			
						
		//Make the results table. Make the strings --> numbers
			setResult("Image Name", resultsCounter, Raw_Names[r]);
			setResult("ROI", resultsCounter, call("ij.plugin.frame.RoiManager.getName", CurrentRoi));
			//setResult("Acquisition Date", resultsCounter, AcqDate);
			setResult("Microscopy Modality", resultsCounter, Modality);
			setResult("Microscope Identifier", resultsCounter, InstrumentIdent);
			setResult("Pinhole [AU]", resultsCounter, Pinhole);
			setResult("Numerical Aperture", resultsCounter, NumAp);
			setResult("Workgroup", resultsCounter, Workgroup);
			setResult("QC Type", resultsCounter, QCType);
			setResult("QC Subtype", resultsCounter, QCSubType);
			setResult("Workgroup", resultsCounter, Workgroup);
			setResult("X Res", resultsCounter, parseFloat(X_l[1]));
			setResult("X R2", resultsCounter, parseFloat(X_l[2]));
			setResult("Y Res", resultsCounter, parseFloat(Y_l[1]));
			setResult("Y R2", resultsCounter, parseFloat(Y_l[2]));
			setResult("Z Res", resultsCounter, parseFloat(Z_l[1]));
			setResult("Z R2", resultsCounter, parseFloat(Z_l[2]));
			setResult("Dataset Name", resultsCounter, Ext.getName("Dataset", DatasetID));
	//		setResult("Project ID", resultsCounter, DatasetID);
			setResult("Project Name", resultsCounter, Ext.getName("Project", ProjectID));
	//		setResult("Dataset ID", resultsCounter, ProjectID);
			updateResults();
			resultsCounter++;
		}
	
		//If there are any analysed beads for the raw file, then the image , with the Rois is imported to Omero
		if(AnalysedBeads){
			nRois=Ext.saveROIs(ImageID, "");
			print("Image " + Raw_Names[r] + ": " + nRois + " ROI(s) saved.");
			print("\n***************************************************************\n");
			Ext.addToTable("FWHM_Results", "Results", ImageID, ""); // The results contents of each raw file are added to the Omero Table
			run("Clear Results");
		}else{
			print("Image " + Raw_Names[r] + ": 0 ROI(s) saved.");
			print("\n***************************************************************\n");
		}
	}
	roiManager("reset");
	run("Close All");
}
//After all Raw files are processed and imported to Omero, then the Results table is saved as an
//Omero table and as a.txt tab-separated file attachement on the Dataset level.
//The summary .pdf is also attached on the Dataset level
//Ext.saveTable("FWHM_Results", "Project", ProjectID);
Ext.saveTable("FWHM_Results", "Dataset", DatasetID);
txt_file = getDir("temp") +DatasetID+"_fwhm_results.csv";
Ext.saveTableAsFile("FWHM_Results", txt_file, ",");
file_id = Ext.addFile("Dataset", DatasetID, txt_file);
deleted = File.delete(txt_file);
summaryPDF_ID=Ext.addFile("Dataset", DatasetID, dir_proc+ExperimentName+datasetName+"_BatchSummary.pdf");
summaryXLS_ID=Ext.addFile("Dataset", DatasetID, dir_proc+ExperimentName+datasetName+"_summary.xls");

Ext.disconnect();
setBatchMode(false);
print("Finished...!");

//////////////////////////////////////////////////////////////////////////////|
////////////////////      FUNCTIONS      /////////////////////////////////////|
//////////////////////////////////////////////////////////////////////////////|

//Returns the Objective NA (new MetroloJ QC 1.3.1.1)
function getPinhole(summaryLines){
	for (i=0;i<summaryLines.length;i++){
		if(startsWith(summaryLines[i], "Pinhole"+cellseparator)){
			tmp=split(summaryLines[i], cellseparator);
			return parseFloat(tmp[1]);
		}
	}
}


//Returns the Objective NA (new MetroloJ QC 1.3.1.1)
function getObjectiveNA(summaryLines){
	for (i=0;i<summaryLines.length;i++){
		if(startsWith(summaryLines[i], "Objective NA"+cellseparator)){
			tmp=split(summaryLines[i], cellseparator);
			return parseFloat(tmp[1]);
		}
	}
}


//Returns the Res, the R-two and the Coordinate of each bead (new MetroloJ QC 1.3.1.1)
function getBeadMeasurements(summaryLines, searchString, b){
	linesFound=newArray(3);
	linesCounter=0;
	for (i=0;i<summaryLines.length;i++){
		if(startsWith(summaryLines[i], searchString+"_bead"+b)){
			linesFound[linesCounter]=summaryLines[i];
			linesCounter++;
		}
	}
	return linesFound;
}


//Returns the no of beads analysed for each image (new MetroloJ QC 1.3.1.1)
function NoOfBeads(summaryLines, RawNameSearch){
	for (i=0;i<summaryLines.length;i++){
		if(startsWith(summaryLines[i], "Files Names"+cellseparator+RawNameSearch+cellseparator)){
			tmp=split(summaryLines[i], cellseparator);
			return parseInt(tmp[2]);
		}
	}
}

//Count how many raw files (.tif or .tiff or .czi or nd2) have been analysed with MetroloJ
//Returns the number
function countRawFiles(dir) {
	countRaw = 0;
   list = getFileList(dir);
   for (i=0; i<list.length; i++) {
       if (endsWith(list[i], ".tif") ||endsWith(list[i], ".tiff")||endsWith(list[i], ".czi")||endsWith(list[i], ".nd2")){
	       countRaw++;
       }
   }
   return countRaw;
}

//Returns the "experiment name": The folder under the Processed folder.
//At the moment it only returns one folder name.
function getExperimentName(dir_proc){
	expir=getFileList(dir_proc);
	return expir[0];
}

//Returns an array with the file paths of the raw files 
function FindRawFilePaths(dir, countRaw) {
	RawFilePaths=newArray(countRaw);
	list = getFileList(dir);
	n=0;
	for (i=0; i<list.length; i++) {
		if (endsWith(list[i], ".tif") ||endsWith(list[i], ".tiff")||endsWith(list[i], ".czi")||endsWith(list[i], ".nd2")){
			RawFilePaths[n] = dir+list[i];
			showProgress(n++, countRaw);
		}
	}
	return RawFilePaths;
}

//Returns an array with the the names of the raw files (without the file-type .tif or .tiff)
function FindRawFileNames(dir, countRaw) {
	Raw_Names=newArray(countRaw);
	list = getFileList(dir);
	n=0;
	for (i=0; i<list.length; i++) {
		if (endsWith(list[i], ".tif") ||endsWith(list[i], ".tiff")||endsWith(list[i], ".czi")||endsWith(list[i], ".nd2")){
			list[i]=replace(list[i], ".tif", "");
			list[i]=replace(list[i], ".tiff", "");
			list[i]=replace(list[i], ".czi", "");
			list[i]=replace(list[i], ".nd2", "");
			Raw_Names[n]=list[i];
			showProgress(n++, countRaw);
		}
	}
	return Raw_Names;
}
//Makes a Roi "around the analysed bead's coordinates. The size of the Roi is 30x30 pixels.
function makeRoi(X, Y, Z, img){
	RoiSize=30;
	RoiSizeHalf=RoiSize/2;
	selectImage(img);
	Stack.setPosition(0, Z, 0);
	makeRectangle(X-RoiSizeHalf, Y-RoiSizeHalf, RoiSize, RoiSize);
}
//////////////////////////////////////////////////////////////////////////////|
//////////////////////////////////////////////////////////////////////////////|
//////////////////////////////////////////////////////////////////////////////|