/**
 * ### Description
 * MimeType converter with Google Apps Script.
 * GitHub: https://github.com/tanaikech/MimeTypeApp
 * 
 * MimeTypeApp v1.0.0
 * 
 * Author: Kanshi Tanaike
 * @class
 */
class MimeTypeApp {

  /**
   * @constructor
   */
  constructor() {
    /** @private */
    this.fileIds = [];

    /** @private */
    this.blobs = [];

    /** @private */
    this.headers = { authorization: "Bearer " + ScriptApp.getOAuthToken() };

    /** @private */
    this.sleep = 2000;
  }

  /**
   * Set file IDs.
   * 
   * @param {Array} fileIds Array including file IDs.
   * @return {MimeTypeApp} Return MimeTypeApp
   *
   */
  setFileIds(fileIds) {
    this.fileIds = fileIds;
    return this;
  }

  /**
   * Set Blobs.
   * 
   * @param {Array} blobs Array including blobs.
   * @return {MimeTypeApp} Return MimeTypeApp
   *
   */
  setBlobs(blobs) {
    this.blobs = blobs;
    return this;
  }

  /**
   * Get conversion list
   * 
   * @return {Object} Return object including the mimeTypes that this script can convert.
   *
   */
  getConversionList() {
    const { importFormats, exportFormats } = this.getFormats_();
    const obj = Object.fromEntries(Object.entries(importFormats).map(([k, [v]]) => [k, exportFormats[v]]));
    return { ...obj, ...exportFormats };
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
  getAs(object) {
    const { mimeType, folderId = "root", onlyCheck = false } = object;

    if (mimeType.includes("application/vnd.google-apps")) {
      console.log(`In the case of "${mimeType}", an array including the file IDs of the converted data are returned as the output.`);
    } else {
      console.log(`In the case of "${mimeType}", an array including the blobs of the converted data are returned as the output.`);
    }

    const dstFolderId = folderId;
    if ((!this.fileIds || this.fileIds.length == 0) && (!this.blobs || this.blobs.length == 0)) {
      throw new Error("Set file IDs or Blobs.");
    }
    if (!mimeType) {
      throw new Error("Set export mimeType.");
    }
    const { importFormats, exportFormats } = this.getFormats_();
    const { routeFileIds, routeBlobs } = this.getRouteObj_({ mimeType, importFormats, exportFormats });

    if (onlyCheck) {
      const checked = [];
      if (routeFileIds.length > 0) {
        checked.push(...routeFileIds.map((e, i) => {
          const temp = { fileId: this.fileIds[i] };
          if (e.length > 0) {
            temp.isPossible = true;
            temp.conversionRoute = e.map(({ from, to }) => from || to);
          } else {
            temp.isPossible = false;
          }
          return temp;
        }));
      }
      if (routeBlobs.length > 0) {
        checked.push(...routeBlobs.map((e, i) => {
          const temp = { blob: this.blobs[i] };
          if (e.length > 0) {
            temp.isPossible = true;
            temp.conversionRoute = e.map(({ from, to }) => from || to);
          } else {
            temp.isPossible = false;
          }
          return temp;
        }));
      }
      return checked;
    }

    if (this.fileIds.length > 0) {
      const r1 = this.fileIds.map((e, i) => {
        const tempFileIds = [];
        const f = DriveApp.getFileById(e);
        const filename = f.getName();
        const srcMimeType = f.getMimeType();
        console.log(`Start: Converting from (${filename}) "${srcMimeType}" to "${mimeType}".`);
        if (routeFileIds[i] == 0) {
          if (srcMimeType == mimeType) {
            console.warn(`End: This conversion is not run because source mimeType and destination mimeType are the same.`);
            return srcMimeType.includes("application/vnd.google-apps") ? e : f.getBlob().copyBlob();
          }
          console.warn(`End: This conversion cannot be achieved.`);
          return null;
        }
        const [, ...r] = routeFileIds[i];
        const res = r.reduce((o, { to, convert }, i, a) => {
          if (convert == "import") {
            o.currentId = this.import_({ srcFileId: o.currentId, filename, dstMimeType: to, dstFolderId });
            if (i != a.length - 1) {
              tempFileIds.push(o.currentId.toString());
            }
          } else if (convert == "export") {
            if (i == a.length - 1) {
              if (mimeType.includes("application/vnd.google-apps")) {
                o.currentId = this.export_({ srcFileId: o.currentId, filename, finalMimeType: to, dstFolderId, returnValue: "fileId" });
              } else {
                o.currentBlob = this.export_({ srcFileId: o.currentId, filename, finalMimeType: to, dstFolderId, returnValue: "blob" });
              }
            } else {
              o.currentId = this.export_({ srcFileId: o.currentId, filename, finalMimeType: to, dstFolderId, returnValue: "fileId" });
              tempFileIds.push(o.currentId.toString());
            }
          }
          return o;
        }, { currentId: e, currentBlob: null });
        tempFileIds.forEach(e => this.deleteTemporalFile_(e));
        console.log(`End: Done.`);
        return mimeType.includes("application/vnd.google-apps") ? res.currentId : res.currentBlob;
      });
      return r1;
    }
    if (this.blobs.length > 0) {
      const r2 = this.blobs.map((e, i) => {
        const tempFileIds = [];
        const filename = e.getName();
        const srcMimeType = e.getContentType();
        console.log(`Start: Converting from (${filename}) "${srcMimeType}" to "${mimeType}".`);
        if (routeBlobs[i] == 0) {
          if (srcMimeType == mimeType) {
            console.warn(`End: This conversion is not run because source mimeType and destination mimeType are the same.`);
            return e;
          }
          console.warn(`End: This conversion cannot be achieved.`);
          return null;
        }
        const [, ...r] = routeBlobs[i];
        const tempId = DriveApp.getFolderById(dstFolderId).createFile(e).getId();
        tempFileIds.push(tempId);
        const res = r.reduce((o, { to, convert }, i, a) => {
          if (convert == "import") {
            o.currentId = this.import_({ srcFileId: o.currentId, filename, dstMimeType: to, dstFolderId });
            if (i != a.length - 1) {
              tempFileIds.push(o.currentId.toString());
            }
          } else if (convert == "export") {
            if (i == a.length - 1) {
              if (mimeType.includes("application/vnd.google-apps")) {
                o.currentId = this.export_({ srcFileId: o.currentId, filename, finalMimeType: to, dstFolderId, returnValue: "fileId" });
              } else {
                o.currentBlob = this.export_({ srcFileId: o.currentId, filename, finalMimeType: to, dstFolderId, returnValue: "blob" });
              }
            } else {
              o.currentId = this.export_({ srcFileId: o.currentId, filename, finalMimeType: to, dstFolderId, returnValue: "fileId" });
              tempFileIds.push(o.currentId.toString());
            }
          }
          return o;
        }, { currentId: tempId, currentBlob: null });
        tempFileIds.forEach(e => this.deleteTemporalFile_(e));
        console.log(`End: Done.`);
        return mimeType.includes("application/vnd.google-apps") ? res.currentId : res.currentBlob;
      });
      return r2;
    }
    return null;
  }

  /**
   * Get thumbnail images as blobs from file IDs.
   * 
   * @param {Number} width Width of the exported thumbnail image. The default is 1000.
   * @param {String} dstFolderId Destination folder ID. The default folder is "root".
   * @return {Array<Object>} Return an array including the thumbnail blobs of the inputted files.
   *
   */
  getThumbnails(width = 1000) {
    if (!this.fileIds || this.fileIds.length == 0) {
      throw new Error("Set file IDs using the setFileIds method. Thumbnails can only be retrieved from file IDs.");
    }
    return this.fileIds.map(id => {
      const file = DriveApp.getFileById(id);
      const filename = file.getName();
      const srcMimeType = file.getMimeType();
      console.log(`Start: Thumbnail image is retrieved from "${filename}" (${srcMimeType}).`);
      const url = `https://drive.google.com/thumbnail?sz=w${width}&id=${id}`;
      let blob = null;
      const res = UrlFetchApp.fetch(url, { headers: this.headers, muteHttpExceptions: true });
      if (res.getResponseCode() == 200) {
        blob = res.getBlob().copyBlob().setName(filename);
      }
      console.log(blob ? "End: Done." : "End: This conversion cannot be achieved.");
      return blob;
    });
  }

  /**
   * Get file ID from files.
   * 
   * @param {Object} object Object for this method.
   * @param {String} object.srcFileId Source file ID.
   * @param {String} object.filename Filename.
   * @param {String} object.dstMimeType Final mimeType.
   * @param {String} object.dstFolderId Destination folder ID.
   * @return {String|Blob} Return file ID or Blob.
   *
   * @private
   */
  import_(object) {
    const { srcFileId, filename, dstMimeType, dstFolderId } = object;
    const url = `https://www.googleapis.com/drive/v3/files/${srcFileId}/copy`;
    const queryObj = { supportsAllDrives: true };
    const payload = { name: filename, mimeType: dstMimeType, parents: [dstFolderId] };
    const request = {
      url: this.addQueryParameters_(url, queryObj),
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      headers: this.headers,
      muteHttpExceptions: true,
    };
    const res = this.fetch_(request);
    const obj = this.fetchErrProc_(res);
    return obj.id;
  }


  /**
   * Get file ID or blob from Google Docs.
   * 
   * @param {Object} object Object for this method.
   * @param {String} object.srcFileId Source file ID.
   * @param {String} object.filename Filename.
   * @param {String} object.finalMimeType Final mimeType.
   * @param {String} object.dstFolderId Destination folder ID.
   * @param {String} object.returnValue Return value. "Blob" or "fileId"
   * @return {String|Blob} Return file ID or Blob.
   *
   * @private
   */
  export_(object) {
    const { srcFileId, filename, finalMimeType, dstFolderId, returnValue = "blob" } = object;
    const queryObj = { supportsAllDrives: true, fields: "exportLinks" };
    const request = {
      url: this.addQueryParameters_(`https://www.googleapis.com/drive/v3/files/${srcFileId}`, queryObj),
      headers: this.headers,
      muteHttpExceptions: true,
    };
    const res = this.fetch_(request);
    const obj = this.fetchErrProc_(res);
    const ress = this.fetch_({ url: obj.exportLinks[finalMimeType], headers: this.headers, muteHttpExceptions: true });
    let rValue = null;
    if (ress.getResponseCode() == 200) {
      const blob = ress.getBlob().copyBlob().setName(filename);
      if (returnValue == "blob") {
        rValue = blob;
      } else if (returnValue == "fileId") {
        rValue = DriveApp.getFolderById(dstFolderId).createFile(blob).getId();
      }
    }
    return rValue;
  }

  /**
   * Get Route object.
   * 
   * @param {Object} object Object including mimeType, exportFormats, importFormats.
   * @return {Array<Object>} Return 2 objects of routeFileIds and routeBlobs.
   *
   * @private
   */
  getRouteObj_(object) {
    const { mimeType, importFormats, exportFormats } = object;
    const route = { routeFileIds: [], routeBlobs: [] };
    if (this.fileIds.length > 0) {
      const r = this.fileIds.map(id => {
        const srcMimeType = DriveApp.getFileById(id).getMimeType();
        return this.getRoute_({ importFormats, exportFormats, srcMimeType, dstMimeType: mimeType });
      });
      route.routeFileIds.push(...r);
    }
    if (this.blobs.length > 0) {
      const r = this.blobs.map(blob => {
        const srcMimeType = blob.getContentType();
        if (!srcMimeType) {
          throw new Error("Please set content type to blob.");
        }
        return this.getRoute_({ importFormats, exportFormats, srcMimeType, dstMimeType: mimeType });
      });
      route.routeBlobs.push(...r);
    }
    return route;
  }


  /**
   * Get Blobs converted to the given mimeType.
   * 
   * @param {Object} object Object including exportFormats, importFormats, srcMimeType, dstMimeType.
   * @return {Array<Object>} Return an array including the route for converting mimeType.
   *
   * @private
   */
  getRoute_(object) {
    const { srcMimeType, dstMimeType } = object;
    if (srcMimeType == dstMimeType) {
      return [];
    }
    let exf = {};
    let imf = {};
    if (object.exportFormats[srcMimeType]) {
      exf = "export";
      imf = "import";
    } else if (object.importFormats[srcMimeType]) {
      exf = "import";
      imf = "export";
    } else {
      return [];
    }
    const route = object[`${exf}Formats`][srcMimeType].reduce((ar, e) => {
      if (e == dstMimeType) {
        ar.push([{ from: srcMimeType }, { to: e, convert: exf }]);
      }
      if (object[`${imf}Formats`][e]) {
        object[`${imf}Formats`][e].forEach(f => {
          if (f == dstMimeType) {
            ar.push([{ from: srcMimeType }, { to: e, convert: exf }, { to: f, convert: imf }]);
          }
          if (object[`${exf}Formats`][f]) {
            object[`${exf}Formats`][f].forEach(g => {
              if (g == dstMimeType) {
                ar.push([{ from: srcMimeType }, { to: e, convert: exf }, { to: f, convert: imf }, { to: g, convert: exf }]);
              }
              if (object[`${imf}Formats`][g]) {
                object[`${imf}Formats`][g].forEach(h => {
                  if (h == dstMimeType) {
                    ar.push([{ from: srcMimeType }, { to: e, convert: exf }, { to: f, convert: imf }, { to: g, convert: exf }, { to: h, convert: imf }]);
                  }
                });
              }
            });
          }
        });
      }
      return ar;
    }, []);
    if (route.length == 0) {
      return route;
    }
    const min = Math.min(...route.map(e => e.length));
    return route.filter(e => e.length == min)[0];
  }

  /**
   * Delete temporal file.
   * 
   * @param {String} fileId File ID.
   * @return {void}
   *
   * @private
   */
  deleteTemporalFile_(fileId) {
    const request = {
      url: `https://www.googleapis.com/drive/v3/files/${fileId}?supportsAllDrives=true`,
      method: "delete",
      headers: this.headers,
      muteHttpExceptions: true,
    };
    const ress = this.fetch_(request);
    this.fetchErrProc_(ress);
  }

  /**
   * Get formats for converting mimeType.
   * 
   * @return {Object} Return object including mimeTypes.
   *
   * @private
   */
  getFormats_() {
    const queryObj = { fields: "importFormats,exportFormats" };
    const request = {
      url: this.addQueryParameters_("https://www.googleapis.com/drive/v3/about", queryObj),
      headers: this.headers,
      muteHttpExceptions: true,
    };
    const res = this.fetch_(request);
    return this.fetchErrProc_(res);
  }

  /**
   * For HTTP request.
   *
   * @param {UrlFetchApp.HTTPResponse}
   * @return {Object}
   * 
   * @private
   */
  fetchErrProc_(object) {
    const text = object.getContentText();
    if ([200, 204].includes(object.getResponseCode())) {
      return text && JSON.parse(text);
    }
    throw new Error(text);
  }

  /**
   * For HTTP request.
   *
   * @param {Object} object Object including request.
   * @return {UrlFetchApp.HTTPResponse}
   * 
   * @private
   */
  fetch_(object) {
    return UrlFetchApp.fetchAll([object])[0];
  }

  /**
   * ### Description
   * This method is used for adding the query parameters to the URL.
   * Ref: https://tanaikech.github.io/2018/07/12/adding-query-parameters-to-url-using-google-apps-script/
   *
   * @param {String} url The base URL for adding the query parameters.
   * @param {Object} obj JSON object including query parameters.
   * @return {String} URL including the query parameters.
   * 
   * @private
   */
  addQueryParameters_(url, obj) {
    if (url === null || obj === null || typeof url != "string") {
      throw new Error(
        "Please give URL (String) and query parameter (JSON object)."
      );
    }
    return (
      (url == "" ? "" : `${url}?`) +
      Object.entries(obj)
        .flatMap(([k, v]) =>
          Array.isArray(v)
            ? v.map((e) => `${k}=${encodeURIComponent(e)}`)
            : `${k}=${encodeURIComponent(v)}`
        )
        .join("&")
    );
  }
}
