/**************
 * Drive helpers
 **************/

function getPromptText() {
  return DocumentApp.openById(PROMPT_DOC_ID).getBody().getText();
}

function getSystemInstructionText() {
  return DocumentApp.openById(SYSTEM_DOC_ID).getBody().getText();
}

function getSchemaJsonText() {
  return DriveApp.getFileById(SCHEMA_FILE_ID).getBlob().getDataAsString();
}


/**************
 * OAuth token (Service Account)
 **************/

function getServiceAccountAccessToken() {
  const raw = PropertiesService.getScriptProperties().getProperty('SERVICE_ACCOUNT_KEY');
  if (!raw) throw new Error('Missing Script Property: SERVICE_ACCOUNT_KEY');

  const sa = JSON.parse(raw);
  const now = Math.floor(Date.now() / 1000);

  const header = { alg: 'RS256', typ: 'JWT' };
  const claims = {
    iss: sa.client_email,
    scope: 'https://www.googleapis.com/auth/cloud-platform https://www.googleapis.com/auth/devstorage.read_write',
    aud: 'https://oauth2.googleapis.com/token',
    iat: now,
    exp: now + 3600
  };

  const encHeader = Utilities.base64EncodeWebSafe(JSON.stringify(header));
  const encClaims = Utilities.base64EncodeWebSafe(JSON.stringify(claims));
  const signingInput = encHeader + '.' + encClaims;

  const signatureBytes = Utilities.computeRsaSha256Signature(signingInput, sa.private_key);
  const encSignature = Utilities.base64EncodeWebSafe(signatureBytes);

  const jwt = signingInput + '.' + encSignature;

  const resp = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: jwt
    }
  });

  const data = JSON.parse(resp.getContentText());
  if (!data.access_token) throw new Error('OAuth token response missing access_token');
  return data.access_token;
}


/**************
 * Gemini call
 **************/

function callGemini(uploadedFiles) {
  if (!uploadedFiles || uploadedFiles.length === 0) {
    throw new Error('No files provided to Gemini.');
  }

  const promptText = getPromptText();
  const systemText = getSystemInstructionText();
  const schemaObj = JSON.parse(getSchemaJsonText());

  const userParts = [
    { text: promptText }
  ];

  uploadedFiles.forEach((f) => {
    userParts.push({
      fileData: {
        fileUri: f.gcsUri,
        mimeType: f.mimeType || 'application/octet-stream'
      }
    });
  });

  const requestBody = {
    systemInstruction: {
      parts: [{ text: systemText }]
    },
    contents: [{
      role: 'user',
      parts: userParts
    }],
    generationConfig: {
      temperature: 0,
      responseMimeType: 'application/json',
      responseSchema: schemaObj
    }
  };

  const resp = UrlFetchApp.fetch(GEMINI_ENDPOINT, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + getServiceAccountAccessToken()
    },
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() >= 300) {
    throw new Error('Gemini call failed: ' + resp.getContentText());
  }

  const raw = JSON.parse(resp.getContentText());
  const text = raw?.candidates?.[0]?.content?.parts?.[0]?.text;

  if (!text) {
    throw new Error('Unexpected Gemini response (missing text).');
  }

  const parsed = JSON.parse(text);
  if (Array.isArray(parsed)) return parsed;
  if (parsed && Array.isArray(parsed.data)) return parsed.data;

  throw new Error('Unexpected JSON shape from model.');
}
