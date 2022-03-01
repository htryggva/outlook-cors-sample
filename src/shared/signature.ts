/* global Office */

import { callApi } from "../helpers/xhr";

export async function fetchDataAndInsertSignature(accessToken: string) {
  const userResponse = await makeUserGraphApiCall(accessToken);
  const orgResponse = await makeOrganizationGraphApiCall(accessToken);
  const signatureData: SignatureData = {
    displayName: userResponse["displayName"] ?? "",
    mail: userResponse["mail"] ?? "",
    jobTitle: userResponse["jobTitle"] ?? "",
    department: userResponse["department"] ?? "",
    company: orgResponse.value[0]["displayName"] ?? "",
  };
  await insertSignature(signatureData);
}

async function makeUserGraphApiCall(accessToken: string) {
  const jsonResponse = await callApi(
    "https://graph.microsoft.com/v1.0/me?$select=displayName,mail,jobTitle,department",
    {
      method: "GET",
      headers: { Authorization: `Bearer ${accessToken}` },
    }
  );

  return JSON.parse(jsonResponse);
}

async function makeOrganizationGraphApiCall(accessToken: string) {
  const jsonResponse = await callApi("https://graph.microsoft.com/v1.0/organization", {
    method: "GET",
    headers: { Authorization: `Bearer ${accessToken}` },
  });

  return JSON.parse(jsonResponse);
}

function insertSignature(signatureData: SignatureData): Promise<void> {
  const userSignature: string = `
    <div style="font-family: Bierstadt, Calibri">
        <div style="font-size: 16px; font-weight: bold">${signatureData.displayName}</div>
        <div>${signatureData.jobTitle}</div>
        <div>${signatureData.department}</div>
        <div>${signatureData.mail}</div>
        <p style="font-size: 16px">${signatureData.company}</p>
    </div>`;

  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.setSignatureAsync(
      userSignature,
      { coercionType: Office.CoercionType.Html },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          resolve(asyncResult.value);
        } else {
          reject(asyncResult.error);
        }
      }
    );
  });
}

type SignatureData = {
  displayName: string;
  mail: string;
  jobTitle: string;
  department: string;
  company: string;
};
