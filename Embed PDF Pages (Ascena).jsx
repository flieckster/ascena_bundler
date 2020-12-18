// Embed PDF Pages - Adobe Photoshop Script
// Description: place PDF pages and TIFF/JPEG files as smart objects into their corresponding source file
// Requirements: Adobe Photoshop CS3, or higher
// Version: 0.14.0, 14/Oct/2020 //@@@
// Author: Trevor Morris (trevor@morris-photographics.com)
// Website: http://morris-photographics.com/
// ============================================================================
// Installation:
// 1. Place script in:
//    PC(32):  C:\Program Files (x86)\Adobe\Adobe Photoshop ##\Presets\Scripts\
//    PC(64):  C:\Program Files\Adobe\Adobe Photoshop ## (64 Bit)\Presets\Scripts\
//    Mac:     <hard drive>/Applications/Adobe Photoshop ##/Presets/Scripts/
// 2. Restart Photoshop
// 3. Choose File > Scripts > Embed PDF Pages
// ============================================================================

// enable double-clicking from Mac Finder or Windows Explorer
#target photoshop

// bring application forward for double-click events
app.bringToFront();

// constants
const SCRIPT_NAME         = 'Embed PDF Pages'; // script name (used for warnings and error messages)
const DEFAULT_BROWSE_PATH = '~'; // default browse path (syntax: '/drive/folder/subfolder/etc')
const LOG_FILE_NAME       = 'Embed PDF Pages Log.txt'; // log file name (including '.txt' extension)
const OPEN_LOG_FILE       = true; // open log file upon completion (true), or do not open log file (false)
const NOTIFY_WHEN_DONE    = true; // show completion notification (true), or do nothing upon completion (false)
const MACOS               = (Folder.fs == 'Macintosh'); // check for Mac/Windows
const PATTERNS = [
	// [pattern name, capture expression]
	['Loft|Ann', /^(\d{5,7}_\d{3,5}(?:_ALT\d|_B\d|_D\d)?)/i],
	['CBK', /^(\d{15}_\d{4})/i],
	['NY&CO', /^(.+)$/i],
	['UNIQLO', /^([a-z]{4}-[0-9a-z]{8}_\d{2}_[0-9a-z]{2})/i],
	['Cacique', /^(cq-\d{6}_\d{10}_\d{7})/i]
];

// global preferences
var prefs = {};

///////////////////////////////////////////////////////////////////////////////
// main - main function
///////////////////////////////////////////////////////////////////////////////
function main() {

	// show dialog; prompt for files and folders
	if (eppDialog().show() != 0) {
		return;
	}

	// get source files
	var re = /\.(?:psd|tiff?|png|jpg?)$/i;
	var sourceFiles = getMatchingFiles(prefs.sourceFolder, 'files', re);
	var sLen = sourceFiles.length;
	if (!sLen) {
		alert('No Source files found in folder:\n' + prefs.sourceFolder.fsName, SCRIPT_NAME, false);
		return;
	}

	// get optional PDF files
	var pdfFiles = [];
	if (prefs.pdfFolder && prefs.pdfFolder.exists) {
		re = /\.pdf$/i;
		pdfFiles = getMatchingFiles(prefs.pdfFolder, 'files', re);
	}

	// get optional TIFF files
	var tiffFiles = [];
	if (prefs.tiffFolder && prefs.tiffFolder.exists) {
		re = /\.tiff?$/i;
		tiffFiles = getMatchingFiles(prefs.tiffFolder, 'files', re);
	}

	// get optional JPEG files
	var jpgFiles = [];
	if (prefs.jpgFolder && prefs.jpgFolder.exists) {
		re = /\.jpe?g$/i;
		jpgFiles = getMatchingFiles(prefs.jpgFolder, 'files', re);
	}

	// check for files to process
	var tLen = tiffFiles.length;
	var jLen = jpgFiles.length;
/*
	if (!pdfFiles.length && !tLen && !jLen) {
		alert('No TIFF, JPEG, or PDF files found.', SCRIPT_NAME, false);
		return;
	}
*/

	// PSD save options
	var saveOptions;
	if (prefs.fileType == 'PSD') {
		saveOptions = new PhotoshopSaveOptions();
		saveOptions.embedColorProfile = true;
		saveOptions.layers = true;
	}
	// TIFF save options
	else {
		saveOptions = new TiffSaveOptions();
		saveOptions.alphaChannels = true;
		saveOptions.embedColorProfile = true;
		saveOptions.imageCompression = TIFFEncoding.TIFFLZW;
		saveOptions.layers = true;
		saveOptions.layerCompression = LayerCompression.ZIP;
		saveOptions.transparency = true;
	}

	// loop variables
	var sourceFile, saveFile;
	var pdfFile, pageCount;
	var doc, docName, keywords;
	var tiffRE, jpgRE;
	var index;
	var textLayer;
	re = /\.\w+$/i;

	// open source file; place PDF pages; save to output folder
	for (var s = 0; s < sLen; s++) {

		sourceFile = sourceFiles[s];
		docName = decodeURI(sourceFile.name).replace(re, '');

		// generate JPEG expression from file name
		// Cacique pattern only: match based on 10-digit swatch number (e.g., "1234567890.jpg")
		if (PATTERNS[prefs.patternIndex][0] == 'Cacique') {
			index = docName.indexOf('_') + 1;
			jpgRE = RegExp('^' + docName.substring(index, index + 10) + '\.jpe?g$', 'i');
		}
		// all other patterns
		else {
			jpgRE = RegExp('^' + escapeRegExp(docName) + '.*\.jpe?g$', 'i');
		}

		// generate RegExp from source file name
		if (tLen) {
			tiffRE = getFileNameExpression(docName);
			if (!tiffRE) {
				log('Error creating regular expression for source file, "' + decodeURI(sourceFile.name) + '".');
				continue;
			}
		}

		// open source file; select top-most layer
		doc = app.open(sourceFile);
		doc.activeLayer = doc.layers[0];

		// get corresponding PDF file (if one exists)
		pdfFile = File(prefs.pdfFolder + '/' + docName + '.pdf');

		// place PDF pages as embedded smart objects
		if (pdfFile.exists) {

			// get PDF page count
			pageCount = getPDFPageCount(pdfFile);

			// place PDF pages as embedded smart objects
			for (var page = 1; page <= pageCount; page++) {

				placePageFromPDF(pdfFile, page);

				if (prefs.pdfLayerName) {
					doc.activeLayer.name = prefs.pdfLayerName + ' (Pg ' + page + ')';
				}
				else {
					doc.activeLayer.name += ' (Pg ' + page + ')';
				}
			}
		}

		// place TIFF as embedded smart objects
		for (var t = 0; t < tLen; t++) {

			if (tiffRE.test(decodeURI(tiffFiles[t].name))) {

				placeFile(tiffFiles[t]);

				if (prefs.tiffLayerName) {
					doc.activeLayer.name = prefs.tiffLayerName;
				}
			}
		}

		// add JPEG as a layer
		for (var j = 0; j < jLen; j++) {

			if (jpgRE.test(jpgFiles[j].name)) {

				layerFromFile(doc, jpgFiles[j]);

				if (prefs.jpgLayerName) {
					doc.activeLayer.name = prefs.jpgLayerName;
				}
			}
		}

		// add text layer
		if (prefs.textContents.length) {
			textLayer = createTextLayer(doc, prefs.textContents, [25, 227, 102], 50, 50);
			centerLayer(doc, textLayer);
		}

		// embed metadata (keywords)
		if (prefs.keywords.length) {
			keywords = doc.info.keywords;
			keywords = (keywords.length && keywords[0].length ? keywords.concat(prefs.keywords) : prefs.keywords);
			doc.info.keywords = keywords;
		}

		// save PSD to output folder
		saveFile = File(prefs.outputFolder + '/' + docName + '.' + prefs.fileType);
		doc.saveAs(saveFile, saveOptions, false, Extension.LOWERCASE);
		doc.close(SaveOptions.DONOTSAVECHANGES);
	}

	// finished
	if (NOTIFY_WHEN_DONE) {
		if (MACOS) {
			app.beep();
		}
		alert('Done!', SCRIPT_NAME, false);
	}

	// open log file
	if (OPEN_LOG_FILE && prefs.logItems) {
		prefs.logFile.execute();
	}
}

///////////////////////////////////////////////////////////////////////////////
// getFileNameExpression - generate RegExp from file name
///////////////////////////////////////////////////////////////////////////////
function getFileNameExpression(fileName) {
	var pattern = PATTERNS[prefs.patternIndex];
	var matches = fileName.match(pattern[1]);
	return (matches ? RegExp('^' + escapeRegExp(fileName), 'i') : null);
}

///////////////////////////////////////////////////////////////////////////////
// escapeRegExp - escape RegExp-reserved characters
// Source: https://stackoverflow.com/questions/3115150/how-to-escape-regular-expression-special-characters-using-javascript
///////////////////////////////////////////////////////////////////////////////
function escapeRegExp(text) {
  return text.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, '\\$&');
}

///////////////////////////////////////////////////////////////////////////////
// eppDialog - get required folders
///////////////////////////////////////////////////////////////////////////////
function eppDialog() {

	// dialog styles
	const PANEL_TOP_MARGIN = 15;
	const BUTTON_WIDTH = 150;
	const LABEL_WIDTH = (MACOS ? 115 : 85);
	const PATH_WIDTH = 400;
	const FIELD_WIDTH = 200;
	const HELP_TIP = 'Optional layer name.\nLeave blank to use file name.';

	// create dialog
	var dialog = new Window('dialog', SCRIPT_NAME, undefined, {closeButton: true});
	dialog.orientation = 'column';
	dialog.alignChildren = 'fill';

	// default values
	prefs.sourceFolder = null;
	prefs.outputFolder = null;
	prefs.pdfFolder = null;
	prefs.tiffFolder = null;
	prefs.jpgFolder = null;
	prefs.fileType = '';
	prefs.pdfLayerName = '';
	prefs.tiffLayerName = '';
	prefs.jpgLayerName = '';
	prefs.keywords = [];

		// source panel
		var sourcePanel = dialog.add('panel');
		sourcePanel.orientation = 'column';
		sourcePanel.text = 'Photoshop Files';
		sourcePanel.margins.top = PANEL_TOP_MARGIN;
		sourcePanel.alignChildren = 'fill';

			// source group
			var sourceGroup = sourcePanel.add('group');
			sourceGroup.orientation = 'row';
//			sourceGroup.alignChildren = 'fill';

				// source folder button
				var sourceButton = sourceGroup.add('button');
				sourceButton.text = 'Choose Folder...';
				sourceButton.preferredSize.width = BUTTON_WIDTH;
				sourceButton.alignment = 'left';
//				sourceButton.active = true;
				sourceButton.onClick = function() {

					// OS folder selection dialog
					var selectedFolder = Folder.selectDialog('Choose source folder:', defaultBrowsePath(prefs.sourceFolder));

					// check selected folder
					if (selectedFolder) {
						sourcePath.text = selectedFolder.fsName;
						prefs.sourceFolder = selectedFolder;
					}
				};

				// source folder path label
				var sourcePath = sourceGroup.add('statictext', undefined, '', {truncate: 'middle'});
				sourcePath.preferredSize.width = PATH_WIDTH;

			// pattern group
			var patternGroup = sourcePanel.add('group');
			patternGroup.orientation = 'row';
//			patternGroup.alignChildren = 'fill';

				// pattern label
				var patternLabel = patternGroup.add('statictext');
				patternLabel.text = 'Filename Pattern:';

				// pattern drop-down
				var patternDD = patternGroup.add('dropdownlist');
				patternDD.onChange = function() {
					prefs.patternIndex = this.selection.index;
				};

					// populate pattern drop-down
					for (var i = 0, len = PATTERNS.length; i < len; i++) {
						patternDD.add('item', PATTERNS[i][0]);
					}

					// default pattern selection
					prefs.patternIndex = 1;
					patternDD.selection = patternDD.items[prefs.patternIndex];

		// PDF panel
		var pdfPanel = dialog.add('panel');
		pdfPanel.orientation = 'column';
		pdfPanel.text = 'PDF Files';
		pdfPanel.margins.top = PANEL_TOP_MARGIN;
		pdfPanel.alignChildren = 'fill';

			// PDF folder group
			var pdfFolderGroup = pdfPanel.add('group');
			pdfFolderGroup.orientation = 'row';

				// PDF folder button
				var pdfButton = pdfFolderGroup.add('button');
				pdfButton.text = 'Choose Folder...';
				pdfButton.preferredSize.width = BUTTON_WIDTH;
				pdfButton.alignment = 'left';
				pdfButton.onClick = function() {

					// OS folder selection dialog
					var selectedFolder = Folder.selectDialog('Choose PDF folder:', defaultBrowsePath(prefs.pdfFolder, prefs.sourceFolder));

					// check selected folder
					if (selectedFolder) {
						pdfPath.text = selectedFolder.fsName;
						prefs.pdfFolder = selectedFolder;
					}
				};

				// PDF folder path label
				var pdfPath = pdfFolderGroup.add('statictext', undefined, '', {truncate: 'middle'});
				pdfPath.preferredSize.width = PATH_WIDTH;

			// PDF layer name group
			var pdfNameGroup = pdfPanel.add('group');
			pdfNameGroup.orientation = 'row';

				// PDF layer name label
				var pdfNameLabel = pdfNameGroup.add('statictext');
				pdfNameLabel.text = 'Layer Name:';
				pdfNameLabel.helpTip = HELP_TIP;

				// PDF layer name field
				var pdfNameField = pdfNameGroup.add('edittext');
				pdfNameField.preferredSize.width = FIELD_WIDTH;
				pdfNameField.helpTip = HELP_TIP;

		// TIFF panel
		var tiffPanel = dialog.add('panel');
		tiffPanel.orientation = 'column';
		tiffPanel.text = 'TIFF Files';
		tiffPanel.margins.top = PANEL_TOP_MARGIN;
		tiffPanel.alignChildren = 'fill';

			// TIFF folder group
			var tiffFolderGroup = tiffPanel.add('group');
			tiffFolderGroup.orientation = 'row';

				// TIFF folder button
				var tiffButton = tiffFolderGroup.add('button');
				tiffButton.text = 'Choose Folder...';
				tiffButton.preferredSize.width = BUTTON_WIDTH;
				tiffButton.alignment = 'left';
				tiffButton.onClick = function() {

					// OS folder selection dialog
					var selectedFolder = Folder.selectDialog('Choose TIFF folder:', defaultBrowsePath(prefs.tiffFolder, prefs.sourceFolder));

					// check selected folder
					if (selectedFolder) {
						tiffPath.text = selectedFolder.fsName;
						prefs.tiffFolder = selectedFolder;
					}
				};

				// TIFF folder path label
				var tiffPath = tiffFolderGroup.add('statictext', undefined, '', {truncate: 'middle'});
				tiffPath.preferredSize.width = PATH_WIDTH;

			// TIFF layer name group
			var tiffNameGroup = tiffPanel.add('group');
			tiffNameGroup.orientation = 'row';

				// TIFF layer name label
				var tiffNameLabel = tiffNameGroup.add('statictext');
				tiffNameLabel.text = 'Layer Name:';
				tiffNameLabel.helpTip = HELP_TIP;

				// TIFF layer name field
				var tiffNameField = tiffNameGroup.add('edittext');
				tiffNameField.preferredSize.width = FIELD_WIDTH;
				tiffNameField.helpTip = HELP_TIP;

		// JPEG panel
		var jpgPanel = dialog.add('panel');
		jpgPanel.orientation = 'column';
		jpgPanel.text = 'JPEG Files';
		jpgPanel.margins.top = PANEL_TOP_MARGIN;
		jpgPanel.alignChildren = 'fill';

			// JPEG folder group
			var jpgFolderGroup = jpgPanel.add('group');
			jpgFolderGroup.orientation = 'row';

				// JPEG folder button
				var jpgButton = jpgFolderGroup.add('button');
				jpgButton.text = 'Choose Folder...';
				jpgButton.preferredSize.width = BUTTON_WIDTH;
				jpgButton.alignment = 'left';
				jpgButton.onClick = function() {

					// OS folder selection dialog
					var selectedFolder = Folder.selectDialog('Choose JPEG folder:', defaultBrowsePath(prefs.jpgFolder, prefs.sourceFolder));

					// check selected folder
					if (selectedFolder) {
						jpgPath.text = selectedFolder.fsName;
						prefs.jpgFolder = selectedFolder;
					}
				};

				// JPEG folder path label
				var jpgPath = jpgFolderGroup.add('statictext', undefined, '', {truncate: 'middle'});
				jpgPath.preferredSize.width = PATH_WIDTH;

			// JPEG layer name group
			var jpgNameGroup = jpgPanel.add('group');
			jpgNameGroup.orientation = 'row';

				// JPEG layer name label
				var jpgNameLabel = jpgNameGroup.add('statictext');
				jpgNameLabel.text = 'Layer Name:';
				jpgNameLabel.helpTip = HELP_TIP;

				// JPEG layer name field
				var jpgNameField = jpgNameGroup.add('edittext');
				jpgNameField.preferredSize.width = FIELD_WIDTH;
				jpgNameField.helpTip = HELP_TIP;

		// output panel
		var outputPanel = dialog.add('panel');
		outputPanel.orientation = 'column';
		outputPanel.text = 'Output Folder';
		outputPanel.margins.top = PANEL_TOP_MARGIN;
		outputPanel.alignChildren = 'fill';

			// output folder group
			var outputGroup = outputPanel.add('group');
			outputGroup.orientation = 'row';

				// output folder button
				var outputButton = outputGroup.add('button');
				outputButton.text = 'Choose Folder...';
				outputButton.preferredSize.width = BUTTON_WIDTH;
				outputButton.alignment = 'left';
				outputButton.onClick = function() {

					// OS folder selection dialog
					var selectedFolder = Folder.selectDialog('Choose output folder:', defaultBrowsePath(prefs.outputFolder, prefs.sourceFolder));

					// check selected folder
					if (selectedFolder) {
						outputPath.text = selectedFolder.fsName;
						prefs.outputFolder = selectedFolder;
					}
				};

				// output folder path label
				var outputPath = outputGroup.add('statictext', undefined, '', {truncate: 'middle'});
				outputPath.preferredSize.width = PATH_WIDTH;

			// file type group
			var fileTypeGroup = outputPanel.add('group');
			fileTypeGroup.orientation = 'row';

				// file type label
				var fileTypeLabel = fileTypeGroup.add('statictext');
				fileTypeLabel.text = 'File Type:';

				// file type dropdown
				var fileTypeDD = fileTypeGroup.add('dropdownlist');
				fileTypeDD.add('item', 'PSD');
				fileTypeDD.add('item', 'TIFF');
				fileTypeDD.items[0].selected = true;

			// keywords group
			var keywordsGroup = outputPanel.add('group');
			keywordsGroup.orientation = 'row';

				// keywords label
				var keywordsLabel = keywordsGroup.add('statictext');
				keywordsLabel.text = 'Keywords:';

				// keywords field
				var keywordsField = keywordsGroup.add('edittext');
				keywordsField.preferredSize.width = FIELD_WIDTH;

			// text layer group
			var textLayerGroup = outputPanel.add('group');
			textLayerGroup.orientation = 'row';

				// textLayer label
				var textLayerLabel = textLayerGroup.add('statictext');
				textLayerLabel.text = 'Text Layer:';

				// textLayer field
				var textLayerField = textLayerGroup.add('edittext');
				textLayerField.preferredSize.width = FIELD_WIDTH;


		// buttons group
		var buttonGroup = dialog.add('group');
		buttonGroup.orientation = 'row';
		buttonGroup.alignChildren = ['right', 'top'];
		buttonGroup.margins.top = 5;

			// Cancel button
			var cancelButton = buttonGroup.add('button');
			cancelButton.text = 'Cancel';
			cancelButton.onClick = function() {
				dialog.close(2);
				app.bringToFront();
			};

			// OK button
			var okButton = buttonGroup.add('button');
			okButton.text = 'OK';
			okButton.onClick = function() {

				// check for required files and folders
				// NOTE: not checking for TIFF/JPEG folder because they're optional
				if (!prefs.sourceFolder) {
					alert('Please choose a Source folder.', SCRIPT_NAME, false);
					return;
				}
				else if (!prefs.outputFolder) {
					alert('Please choose an Output folder.', SCRIPT_NAME, false);
					return;
				}
/*
				// must have a PDF/TIFF/JPEG folder
				if (!prefs.pdfFolder && !prefs.tiffFolder && !prefs.jpgFolder) {
					alert('Please choose a PDF, TIFF, or JPEG folder.', SCRIPT_NAME, false);
					return;
				}
*/
				// get file type / extension
				prefs.fileType = fileTypeDD.items[fileTypeDD.selection.index].text;

				// get optional layer names
				prefs.pdfLayerName = removeSpaces(pdfNameField.text);
				prefs.tiffLayerName = removeSpaces(tiffNameField.text);
				prefs.jpgLayerName = removeSpaces(jpgNameField.text);
				prefs.keywords = [String(keywordsField.text)];
				prefs.textContents = String(textLayerField.text);

				// close dialog
				dialog.close(0);
				app.bringToFront();
			};

	// set dialog properties
	dialog.defaultElement = okButton;
	dialog.cancelElement = cancelButton;
	dialog.center();

	return dialog;
}

///////////////////////////////////////////////////////////////////////////////
// defaultBrowsePath - get the default browse/folder location
///////////////////////////////////////////////////////////////////////////////
function defaultBrowsePath(folder, backupFolder) {

	// 1. attempt to get the "current" folder (if one exists)
	if (folder && folder.exists) {
		return folder;
	}
	// 2. attempt to get the "current" folder's parent
	else if (folder && folder.parent.exists) {
		return folder.parent;
	}
	// 3. if no "current" folder exists, try a "backup" folder
	// (e.g., the parent of the "Source" folder)
	else if (backupFolder && backupFolder.exists) {
		return backupFolder.parent;
	}
	// 4. if all else fails, return the "default" folder
	else {
		return Folder(DEFAULT_BROWSE_PATH);
	}
}

///////////////////////////////////////////////////////////////////////////////
// removeSpaces - remove leading/trailing spaces
///////////////////////////////////////////////////////////////////////////////
function removeSpaces(text) {
	text = text.replace(/^\s+/, '');
	text = text.replace(/\s+$/, '');
	return (text.length ? text : undefined);
}

///////////////////////////////////////////////////////////////////////////////
// getPDFPageCount - get the page count for the referenced PDF file
///////////////////////////////////////////////////////////////////////////////
function getPDFPageCount(file) {

	// set maximum page count
	var maxPages = 100;

	// PDF open options
	var openOptions = new PDFOpenOptions();
	openOptions.antiAlias = true;
	openOptions.bitsPerChannel = BitsPerChannelType.EIGHT;
	openOptions.constrainProportions = true;
	openOptions.cropPage = CropToType.MEDIABOX; // CropToType.BLEEDBOX
	openOptions.mode = OpenDocumentMode.RGB; // OpenDocumentMode.CMYK
	openOptions.resolution = 72;
	openOptions.suppressWarnings = true;

	// loop variables
	var doc;

	// open the PDF one page at a time
	while (maxPages > 0) {

		// set the page number
		openOptions.page = maxPages;

		// attempt to open the page; return page count
		try {
			doc = open(file, openOptions);
			doc.close(SaveOptions.DONOTSAVECHANGES);
			return maxPages;
		}
		// page is out of range
		catch(e) {
			--maxPages;
			continue;
		}
	}

	// default to one page, just in case
	return 1;
}

///////////////////////////////////////////////////////////////////////////////
// placePageFromPDF - place the specified PDf page as an embedded smart object
// NOTE: unfortunately, this function will not error when the page number is
// out of range (i.e., doesn't exist); instead, it will place the last page
///////////////////////////////////////////////////////////////////////////////
function placePageFromPDF(pdfFile, page) {

	var desc1 = new ActionDescriptor();
	var desc2 = new ActionDescriptor();
	desc2.putEnumerated(cTID('fsel'), sTID('pdfSelection'), sTID('page'));
	desc2.putInteger(cTID('PgNm'), page);
	desc2.putEnumerated(cTID('Crop'), sTID('cropTo'), sTID('boundingBox'));
	desc2.putBoolean(sTID('suppressWarnings'), false);
	desc2.putBoolean(cTID('AntA'), true);
	desc2.putBoolean(cTID('ClPt'), true);
	desc1.putObject(cTID('As  '), cTID('PDFG'), desc2);
	desc1.putInteger(cTID('Idnt'), 4);
	desc1.putPath(cTID('null'), pdfFile);
	desc1.putEnumerated(cTID('FTcs'), cTID('QCSt'), cTID('Qcsa'));
	var desc3 = new ActionDescriptor();
	desc3.putUnitDouble(cTID('Hrzn'), cTID('#Pxl'), 0);
	desc3.putUnitDouble(cTID('Vrtc'), cTID('#Pxl'), 0);
	desc1.putObject(cTID('Ofst'), cTID('Ofst'), desc3);
	desc1.putBoolean(cTID('AntA'), true);

	try {
		app.executeAction(cTID('Plc '), desc1, DialogModes.NO);
		return true;
	}
	catch(e) {
		return false;
	}
}

///////////////////////////////////////////////////////////////////////////////
// placeFile - File > Place Embedded
///////////////////////////////////////////////////////////////////////////////
function placeFile(file) {

	var desc1 = new ActionDescriptor();
	desc1.putPath(cTID('null'), file);
	desc1.putEnumerated(cTID('FTcs'), cTID('QCSt'), cTID('Qcsa'));

	try {
		app.executeAction(cTID('Plc '), desc1, DialogModes.NO);
		return true;
	}
	catch(e) {
		return false;
	}
}

///////////////////////////////////////////////////////////////////////////////
// layerFromFile - add the reference file as a layer within the provided document
///////////////////////////////////////////////////////////////////////////////
function layerFromFile(doc, file) {

	// open image
	var img = null;
	try {
		img = app.open(file);
	}
	catch(e) {
		log('Error opening image file, "' + decodeURI(file.name) + '".');
		return false;
	}

	// duplicate image to document
	img.activeLayer.duplicate(doc, ElementPlacement.PLACEATBEGINNING);

	// close image
	img.close(SaveOptions.DONOTSAVECHANGES);

	// rename layer
	var re = /\.\w+$/;
	doc.activeLayer.name = decodeURI(file.name).replace(re, '');

	return true;
}

///////////////////////////////////////////////////////////////////////////////
// centerLayer - center the provided layer in the active document
///////////////////////////////////////////////////////////////////////////////
function centerLayer(doc, layer) {

	// layer properties
	var bounds = layer.bounds;
	var left = Number(bounds[0]);
	var top = Number(bounds[1]);
	var layerWidth = Number(bounds[2]) - left;
	var layerHeight = Number(bounds[3]) - top;

	// check for empty layer
	if (!layerWidth) {
		return;
	}

	// document properties
	var docWidth = Number(doc.width);
	var docHeight = Number(doc.height);

	// center layer
	layer.translate(
		(docWidth - layerWidth) / 2 - left,
		(docHeight - layerHeight) / 2 - top
	);
}

///////////////////////////////////////////////////////////////////////////////
// createTextLayer - create text layer
///////////////////////////////////////////////////////////////////////////////
function createTextLayer(doc, contents, color, x, y) {

	// action descriptor
	var desc1 = new ActionDescriptor();
	var desc2 = new ActionDescriptor();
	desc2.putString(cTID('Txt '), contents);

	// layer reference
	var ref = new ActionReference();
	ref.putClass(cTID('TxLr'));
	desc1.putReference(cTID('null'), ref);

	// position (percentage)
	var desc4 = new ActionDescriptor();
	desc4.putUnitDouble(cTID('Hrzn'), cTID('#Prc'), x);
	desc4.putUnitDouble(cTID('Vrtc'), cTID('#Prc'), y);
	desc2.putObject(cTID('TxtC'), cTID('Pnt '), desc4);

	desc2.putEnumerated(sTID('textGridding'), sTID('textGridding'), cTID('None'));
	desc2.putEnumerated(cTID('Ornt'), cTID('Ornt'), cTID('Hrzn'));
	desc2.putEnumerated(cTID('AntA'), cTID('Annt'), sTID('antiAliasSharp'));

	var list1 = new ActionList();

	var desc5 = new ActionDescriptor();
	desc5.putInteger(cTID('From'), 0);
	desc5.putInteger(cTID('T   '), cTID('null')); // cover full range of text

	// font properties
	var desc6 = new ActionDescriptor();
	desc6.putString(sTID('fontPostScriptName'), 'Verdana');
	desc6.putString(cTID('FntN'), 'Verdana');
	desc6.putString(cTID('FntS'), 'Regular');
	desc6.putInteger(cTID('Scrp'), 0);
	desc6.putInteger(cTID('FntT'), 1);
	desc6.putUnitDouble(cTID('Sz  '), cTID('#Pnt'), 50);
	desc6.putDouble(cTID('HrzS'), 100);
	desc6.putDouble(cTID('VrtS'), 100);

/*
	// additional font properties
	desc6.putBoolean(sTID('syntheticBold'), false);
	desc6.putBoolean(sTID('syntheticItalic'), false);
	desc6.putBoolean(sTID('autoLeading'), true);
	desc6.putInteger(cTID('Trck'), 0);
	desc6.putUnitDouble(cTID('Bsln'), cTID('#Pnt'), 0);
	desc6.putDouble(sTID('characterRotation'), 0);
	desc6.putEnumerated(cTID('AtKr'), cTID('AtKr'), sTID('metricsKern'));
	desc6.putEnumerated(sTID('fontCaps'), sTID('fontCaps'), cTID('Nrml'));
	desc6.putEnumerated(sTID('baseline'), sTID('baseline'), cTID('Nrml'));
	desc6.putEnumerated(sTID('otbaseline'), sTID('otbaseline'), cTID('Nrml'));
	desc6.putEnumerated(sTID('strikethrough'), sTID('strikethrough'), sTID('strikethroughOff'));
	desc6.putEnumerated(cTID('Undl'), cTID('Undl'), sTID('underlineOff'));
	desc6.putUnitDouble(sTID('underlineOffset'), cTID('#Pnt'), 0);
*/

	// color
	var desc7 = new ActionDescriptor();
	desc7.putDouble(cTID('Rd  '), color[0]);
	desc7.putDouble(cTID('Grn '), color[1]);
	desc7.putDouble(cTID('Bl  '), color[2]);
	desc6.putObject(cTID('Clr '), cTID('RGBC'), desc7);

	desc5.putObject(cTID('TxtS'), cTID('TxtS'), desc6);
	list1.putObject(cTID('Txtt'), desc5);
	desc2.putList(cTID('Txtt'), list1);
/*
	// begin alignment
	var list2 = new ActionList();
	var desc8 = new ActionDescriptor();
	desc8.putInteger(cTID('From'), 0);
	desc8.putInteger(cTID('T   '), cTID('null')); // cover full range of text

	var desc9 = new ActionDescriptor();
	desc9.putEnumerated(cTID('Algn'), cTID('Alg '), cTID(alignment));

	desc8.putObject(sTID('paragraphStyle'), sTID('paragraphStyle'), desc9);
	list2.putObject(sTID('paragraphStyleRange'), desc8);
	desc2.putList(sTID('paragraphStyleRange'), list2);
	// end alignment
*/
	desc1.putObject(cTID('Usng'), cTID('TxLr'), desc2);

	// create text layer
	try {
		app.executeAction(cTID('Mk  '), desc1, DialogModes.NO);
	}
	catch(e) {
		alert(e);
	}

	return doc.activeLayer;
}

///////////////////////////////////////////////////////////////////////////////
// getMatchingFiles - get matching files and/or folders, with optional pattern
///////////////////////////////////////////////////////////////////////////////
function getMatchingFiles(folder, kind, pattern, recursive) {

	// check if folder exists
	folder = new Folder(folder);
	if (!folder.exists) {
		alert(
			'Folder not found:\r' + decodeURI(folder.name),
			SCRIPT_NAME,
			true
		);
		return [];
	}

	// ignore system files and hidden files
	var ignoreRE = /^(?:~|\.)/;

	// match files, folders, or both (default)
	var matchFiles = (/^folder/i.test(kind) ? false : true);
	var matchFolders = (/^file/i.test(kind) ? false : true);

	// get matching files and/or folders
	var files = [];
	doGetMatchingFiles(folder);
	return files;


	// get matching files and/or folders, recursively
	function doGetMatchingFiles(folder) {

		// local variables
		var allFiles = folder.getFiles();
		var file, fileName;

		// test all files
		for (var i = 0, len = allFiles.length; i < len; i++) {

			file = allFiles[i];
			fileName = decodeURI(file.name);

			// skip system files and hidden files
			if (ignoreRE.test(fileName)) {
				continue;
			}

			// check files
			if (matchFiles && file instanceof File) {
				if (!pattern || (pattern && pattern.test(fileName))) {
					files.push(file);
				}
			}

			// check folders
			if (matchFolders && file instanceof Folder) {
				if (!pattern || (pattern && pattern.test(fileName))) {
					files.push(file);
				}
			}

			// get files in subfolders
			if (recursive && file instanceof Folder) {
				doGetMatchingFiles(file);
			}
		}
	}
}

///////////////////////////////////////////////////////////////////////////////
// initLogFile - initialize log file
///////////////////////////////////////////////////////////////////////////////
function initLogFile(folder) {

	// initialize log file
	var logFile = File(folder + '/' + LOG_FILE_NAME);
	logFile.open('w');
	logFile.writeln(SCRIPT_NAME + ' Script Log');
	logFile.writeln('OS: ' + $.os);
	logFile.writeln(app.name + ' ' + app.version);
	logFile.writeln();
	logFile.close();

	// update global preferences
	prefs.logFile = logFile;
	prefs.logItems = 0;
}

///////////////////////////////////////////////////////////////////////////////
// log - write to log file with error details
///////////////////////////////////////////////////////////////////////////////
function log(msg, e) {

	// check if logging is disabled
	if (prefs.logDisabled) {
		return;
	}

	// init log file
	if (!prefs.logFile) {
		initLogFile(prefs.outputFolder.path);
	}

	// open log file; jump to end of file
	var file = prefs.logFile;
	file.open('e');
	file.seek(0, 2);

	// handle Mac line endings
	if (MACOS) {
		file.lineFeed = 'Unix';
	}

	// get ISO date string
	var dateTime = toISODateString() + ' - ';

	// get Photoshop error (if available)
	var error = '';
	if (e) {
		error = '\n' + e + ' @ ' + e.line;
	}

	// write error to log
	file.writeln(dateTime + msg + error);

	// close log file
	file.close();

	// update log count
	prefs.logItems += 1;
}

///////////////////////////////////////////////////////////////////////////////
// toISODateString - format a Date object into a proper ISO 8601 date string
// Copyright: (c)2017, xbytor
// License: http://creativecommons.org/licenses/LGPL/2.1
// Contact: xbytor@gmail.com
///////////////////////////////////////////////////////////////////////////////
function toISODateString(date, timeDesignator, dateOnly, precision) {

	if (!date) date = new Date();
	if (timeDesignator == undefined) {timeDesignator = ' T';}

	var str = '';
	var ms;

	if (date instanceof Date) {

		str = (date.getFullYear() + '-' +
		zeroPad(date.getMonth() + 1, 2) + '-' +
		zeroPad(date.getDate(), 2));

		if (!dateOnly) {

			str += (timeDesignator +
			zeroPad(date.getHours(), 2) + ':' +
			zeroPad(date.getMinutes(), 2) + ':' +
			zeroPad(date.getSeconds(), 2));

			if (precision && typeof(precision) == 'number') {

				ms = date.getMilliseconds();

				if (ms) {
					var millis = zeroPad(ms.toString(), precision);
					var s = millis.slice(0, Math.min(precision, millis.length));
					str += '.' + s;
				}
			}
		}
	}

	return str;
}

///////////////////////////////////////////////////////////////////////////////
// zeroPad - return zero-padded value
///////////////////////////////////////////////////////////////////////////////
function zeroPad(num, digits) {

	num = String(num || 0);
	digits = (digits || 2);

	var zeros = String(Math.pow(10, digits - num.length)).substr(1);
	return zeros + num;
}

///////////////////////////////////////////////////////////////////////////////
// isRequiredVersion - check for the required version of Adobe Photoshop
///////////////////////////////////////////////////////////////////////////////
function isRequiredVersion() {
	if (parseInt(app.version) >= 10) {
		return true;
	}
	else {
		alert(
			'This script requires Adobe Photoshop CS3 or higher.',
			SCRIPT_NAME,
			false
		);
		return false;
	}
}

///////////////////////////////////////////////////////////////////////////////
// showError - display error message if something goes wrong
///////////////////////////////////////////////////////////////////////////////
function showError(error) {
	alert(
		error + ': on line ' + error.line,
		SCRIPT_NAME,
		true
	);
}

///////////////////////////////////////////////////////////////////////////////
// cTID - alias the native app.charIDToTypeID function
// Credit: adapted from xtools <http://ps-scripts.sourceforge.net/xtools.html>
///////////////////////////////////////////////////////////////////////////////
function cTID(s) {
	if (!cTID[s]) {
		cTID[s] = app.charIDToTypeID(s);
	}
	return cTID[s];
}

///////////////////////////////////////////////////////////////////////////////
// sTID - alias the native app.stringIDToTypeID function
// Credit: adapted from xtools <http://ps-scripts.sourceforge.net/xtools.html>
///////////////////////////////////////////////////////////////////////////////
function sTID(s) {
	if (!sTID[s]) {
		sTID[s] = app.stringIDToTypeID(s);
	}
	return sTID[s];
}


// test initial conditions prior to running main function
if (isRequiredVersion()) {

	// suppress open options dialogs
	var dialogs = app.displayDialogs;
	app.displayDialogs = DialogModes.NO;

	// remember unit settings; switch to pixels
	var rulerUnits = app.preferences.rulerUnits;
	app.preferences.rulerUnits = Units.PIXELS;

	// suspend history to speed things up
	try {
		main();
	}
	// don't report error on user cancel
	catch(e) {
		if (e.number != 8007) {
			showError(e);
		}
	}

	// restore original unit setting
	app.preferences.rulerUnits = rulerUnits;

	// restore dialogs
	app.displayDialogs = dialogs;
}
