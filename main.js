/**
  This is used when MimeTypeApp is used as a Google Apps Script library.
*/

/**
 * Set file IDs.
 * 
 * @param {Array} fileIds Array including file IDs.
 * @return {MimeTypeApp} Return MimeTypeApp.
 *
 */
function setFileIds(fileIds) {
  return new MimeTypeApp().setFileIds(fileIds);
}

/**
 * Set Blobs.
 * 
 * @param {Array} blobs Array including blobs.
 * @return {MimeTypeApp} Return MimeTypeApp.
 *
 */
function setBlobs(blobs) {
  return new MimeTypeApp().setBlobs(blobs);
}

/**
 * Get conversion list
 * 
 * @return {Object} Return object including the mimeTypes that this script can convert.
 *
 */
function getConversionList() {
  return new MimeTypeApp().getConversionList();
}

/**
 * Get file IDs converted to the given mimeType.
 * 
 * @param {Object} object Object including setting data.
 * @param {String} object.mimeType Target mimeType.
 * @param {String} object.folderId Destination folder ID. The default folder is "root." When the target MIME type is a Google Docs file (Document, Spreadsheet, Slide, etc.), the converted data is required to be created as a file. When this is set, the created files are put into this folder.
 * @param {Boolean} object.onlyCheck The default is false. When set to true, it verifies whether the input file IDs or blobs can be converted to the target MIME type and returns the verified result.
 * @return {Array<Object>} Return an array containing the blobs or file IDs of the converted files. When the target MIME type is a Google Docs file (Document, Spreadsheet, Slide, etc.), the destination data type must be a file. Therefore, the file ID is returned. When the target MIME type is not a Google Docs file, the blob can be returned as the destination data type. Hence, the blob is returned.
 */
function getAs(object) {
  return new MimeTypeApp().getAs(object);
}

/**
 * Get thumbnail images as blobs from file IDs.
 * 
 * @param {Number} width Width of the exported thumbnail image. The default is 1000.
 * @param {String} dstFolderId Destination folder ID. The default folder is "root".
 * @return {Array<Object>} Return an array including the thumbnail blobs of the inputted files.
 *
 */
function getThumbnails(width = 1000) {
  new MimeTypeApp().getThumbnails(width);
}
