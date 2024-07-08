// @File (label="Choose a root Directory with raw images ", style="directory") dir
// @String(label="Host", value='wss://workshop.openmicroscopy.org/omero-ws') omrsrv
// @Integer(label="Port", value=4046) omrport
// @String(label="Omero Username", style="Text Field") omrusr
// @String(label="Password", style='password', persist=false) omrpwd
// @String(label="Omero Project", style="Text Field") omrProject
// @Boolean(label="Batch Mode", value=false) batch

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
 *****************************************************************************
 * Description of macro
 * --------------------
 * This macro is used for uploading on an omero server the results of the PSF 
 * analysis with MetroloJ QC. It uploads the raw files analysed by MetroloJ with rois 
 * of the analysed beads. It also generates a table (attached to the dataset level) with
 * the FWHM measurements as well as the R2 for each analysed bead.
 * It also uploads (attaches) the results as a tab-separated .txt file and the summary.pdf file 
 * at the Dataset level.
 * 
 * Running the script:
 */
requires("1.54f");
//			 Dependencies
// Tested with the latest versions of plugings as mentioned in:
// https://github.com/GReD-Clermont/simple-omero-client
// and
// https://github.com/GReD-Clermont/omero_macro-extensions
// as well as the latest version of Omero.insight plugin:
//https://www.openmicroscopy.org/omero/downloads/
//
run("OMERO Extensions");
sep=File.separator;
dir=dir+sep;
//dir = getDirectory("Choose a root Directory with raw images ");
dir_proc=dir +"Processed/";
if(batch){
	setBatchMode(true);
}
countRaw=countRawFiles(dir);
RawFilePaths=FindRawFilePaths(dir, countRaw);
Raw_Names=FindRawFileNames(dir, countRaw);
ExperimentName=getExperimentName(dir_proc);
datasetName=replace(ExperimentName, "/", "");
lineseparator = "\n";
summaryLines=split(File.openAsString(dir_proc+ExperimentName+"summary.xls"), lineseparator);

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
if(oldProject!=""){
	ProjectID=oldProject;
	print("Project with ID "+ProjectID+" and name "+Ext.getName("Project", ProjectID)+" already exists and will be used.");
}else{
	ProjectID=Ext.createProject(omrProject, "Project for Uploading MetroloJ QC PSF Results");
	print("New project with ID "+ProjectID+" and name "+Ext.getName("Project", ProjectID)+" created.");
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
print("");
print("");
print("Start with raw files....");
run("Clear Results");
AnalysedBeads=false;
for (r=0; r<RawFilePaths.length; r++){
	open(RawFilePaths[r]);
	img=getImageID();
	roiManager("reset");
	run("Clear Results");
	csvPath=dir_proc+ExperimentName+Raw_Names[r]+sep+"beadCoordinates.xls";
	
	lineseparator = "\n";
	cellseparator = "\t";

     // Finds Coordinates of beads and makes rois
	lines=split(File.openAsString(csvPath), lineseparator);
	resultsCounter=0;
	for (i=0;i<lines.length; i++){
		columns=split(lines[i], cellseparator);
		X=columns[1];
		Y=columns[2];
		Z=columns[3];
		status=columns[4];
		beadNo=parseInt(columns[5]);
		bead=columns[6];
		if(status == "Analysed"){
			AnalysedBeads=true;
			run("Select None");
			makeRoi(X, Y, Z, img);
			roiManager("add");
			CurrentRoi=roiManager("count")-1;
			roiManager("select", CurrentRoi);
			roiManager("rename", Raw_Names[r]+"_"+bead);
			Res="";
			R="";
//			print (Raw_Names[r]);
			for(q=0;q<summaryLines.length;q++){
				searchString=Raw_Names[r];					//MetroloJQC Processes the Image Names and removes spaces (probably) at the summary.xls
				searchString=replace(searchString, " - ", "-");
				if(startsWith(summaryLines[q], searchString+"_"+bead)){
					colSum=split(summaryLines[q], cellseparator);
					Res=Res+"&"+colSum[1];
					R=R+"&"+colSum[2];
				}
			}
			Resolution=split(Res, "&");
			R_two=split(R, "&");
			setResult("Label", resultsCounter, Raw_Names[r]); //+"_"+bead
			setResult("ROI", resultsCounter, bead);
			setResult("X Res", resultsCounter, parseFloat(Resolution[0]));
			setResult("X R2", resultsCounter, parseFloat(R_two[0]));
			setResult("Y Res", resultsCounter, parseFloat(Resolution[1]));
			setResult("Y R2", resultsCounter, parseFloat(R_two[1]));
			setResult("Z Res", resultsCounter, parseFloat(Resolution[2]));
			setResult("Z R2", resultsCounter, parseFloat(R_two[2]));
			updateResults();
			resultsCounter++;
		}
	}
	//Here goes the upload to OMERO: For each image upload the tif file with Rois and results as tables (maybe fix about tags from file name?)
	if(AnalysedBeads){
		selectImage(img);
		rename(Raw_Names[r]);
		imageID=Ext.importImage(DatasetID);
		print("creating summary results for image ID " + imageID + ": " +Raw_Names[r]);
		nRois=Ext.saveROIs(imageID, "");
		print("Image " + Raw_Names[r] + ": " + nRois + " ROI(s) saved.");
		print("***************************************************************");
		Ext.addToTable("FWHM_Results", "Results", imageID);
		run("Clear Results");
		
	}
	roiManager("reset");
	run("Close All");
}
Ext.saveTable("FWHM_Results", "Dataset", DatasetID);
txt_file = getDir("temp") +DatasetID+"_fwhm_results.txt";
Ext.saveTableAsFile("FWHM_Results", txt_file, ",");
file_id = Ext.addFile("Dataset", DatasetID, txt_file);
deleted = File.delete(txt_file);
summaryPDF_ID=Ext.addFile("Dataset", DatasetID, dir_proc+ExperimentName+"summary.pdf");

Ext.disconnect();
setBatchMode(false);
print("Finished...!");


// FUNCTIONS

//Count how many raw files (.tif or .tiff) have been analysed with MetroloJ
//Returns the number

function countRawFiles(dir) {
	countRaw = 0;
   list = getFileList(dir);
   for (i=0; i<list.length; i++) {
       if (endsWith(list[i], ".tif") ||endsWith(list[i], ".tiff")){
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
			Raw_Names[n]=list[i];
			showProgress(n++, countRaw);
		}
	}
	return Raw_Names;
}
/*
function processFile(path) {
    if (endsWith(path, ".tif") ||endsWith(path, ".tiff") || endsWith(path, ".jpg")) {
        print("Found Image File : " + path);
   }else if(endsWith(path, ".xls")){
   	print("Found CSV File : " + path);
   }
}
*/
function makeRoi(X, Y, Z, img){
	selectImage(img);
	Stack.setPosition(0, Z, 0);
	makeRectangle(X-15, Y-15, 30, 30);
}
