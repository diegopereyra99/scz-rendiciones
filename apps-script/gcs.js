function uploadBlobToGCS_(bucket, objectPath, blob) {
  const url =
    'https://storage.googleapis.com/upload/storage/v1/b/' +
    encodeURIComponent(bucket) +
    '/o?uploadType=media&name=' +
    encodeURIComponent(objectPath);

  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: blob.getContentType() || 'application/zip',
    headers: {
      Authorization: 'Bearer ' + getServiceAccountAccessToken()
    },
    payload: blob.getBytes(),
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  if (code >= 300) {
    throw new Error('GCS upload failed (' + code + '): ' + resp.getContentText());
  }

  return `gs://${bucket}/${objectPath}`;
}


function uploadDriveFolderToGCS_(driveFolderId, bucket, gcsBasePath) {
  const folder = DriveApp.getFolderById(driveFolderId);
  const files = folder.getFiles();

  const uploaded = [];

  while (files.hasNext()) {
    const file = files.next();

    // Ignorar Google Docs / Sheets / Slides
    if (file.getMimeType().startsWith('application/vnd.google-apps')) {
      continue;
    }

    const blob = file.getBlob();
    const objectPath = `${gcsBasePath}/${file.getName()}`;

    uploadBlobToGCS_(bucket, objectPath, blob);

    uploaded.push({
      name: file.getName(),
      mimeType: blob.getContentType(),
      gcsUri: `gs://${bucket}/${objectPath}`
    });

    console.log('Uploaded:', objectPath);
  }

  if (uploaded.length === 0) {
    throw new Error('Folder is empty or only contains Google Docs');
  }

  return uploaded;
}

function parseGcsUri_(gcsUri) {
  if (!gcsUri || !gcsUri.startsWith('gs://')) {
    throw new Error(`Invalid gcsUri: ${gcsUri}`);
  }
  const noScheme = gcsUri.slice('gs://'.length);
  const slash = noScheme.indexOf('/');
  if (slash < 0) throw new Error(`Invalid gcsUri (missing object path): ${gcsUri}`);
  const bucket = noScheme.slice(0, slash);
  const object = noScheme.slice(slash + 1);
  return { bucket, object };
}


function downloadGcsToBlob_(gcsUri, filenameOpt) {
  const { bucket, object } = parseGcsUri_(gcsUri);

  // Endpoint: GET .../b/<bucket>/o/<urlencoded object>?alt=media
  const url =
    `https://storage.googleapis.com/storage/v1/b/${encodeURIComponent(bucket)}` +
    `/o/${encodeURIComponent(object)}?alt=media`;

  const token = getServiceAccountAccessToken(); // <- tu función existente
  const resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true,
    followRedirects: true,
  });

  const code = resp.getResponseCode();
  if (code !== 200) {
    throw new Error(`GCS download failed (${code}): ${resp.getContentText()}`);
  }

  const blob = resp.getBlob();
  if (filenameOpt) blob.setName(filenameOpt);
  return blob;
}



function saveGcsFileToDrive_(gcsUri, filename, folderId) {
  const blob = downloadGcsToBlob_(gcsUri); // <- ya lo tenés o lo implementás
  blob.setName(filename);
  return DriveApp.getFolderById(folderId).createFile(blob);
}



