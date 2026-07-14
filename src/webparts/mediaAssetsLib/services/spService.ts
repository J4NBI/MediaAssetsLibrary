import { SPHttpClient } from "@microsoft/sp-http";
/**
 * Holt das FormDigestValue von SharePoint für POST/MERGE-Operationen
 * Erforderlich für Schreibzugriff auf SharePoint REST API
 * @private
 * @async
 * @returns {Promise<string>} Das FormDigestValue Token
 * @throws {Error} Wenn der Digest-Abruf fehlschlägt
 */
export const getRequestDigest = async (
  siteUrl: string,
  spHttpClient: SPHttpClient,
): Promise<string> => {
  const response = await spHttpClient.post(
    `${siteUrl}/_api/contextinfo`,
    SPHttpClient.configurations.v1,
    {
      headers: {
        Accept: "application/json;odata=nometadata",
      },
    },
  );

  const data = await response.json();

  return data.FormDigestValue;
};

/**
 * Ruft Informationen des aktuellen angemeldeten Benutzers ab
 * Wird verwendet, um den Ersteller bei Upload zu speichern
 * @private
 * @async
 * @returns {Promise<{id: number, title: string}>} Benutzer-ID und Name
 * @throws {Error} Wenn der Benutzerabruf fehlschlägt
 */

export const getCurrentUser = async (
  siteUrl: string,
  spHttpClient: SPHttpClient,
): Promise<{ id: number; title: string }> => {
  const response = await spHttpClient.get(
    `${siteUrl}/_api/web/currentuser`,
    SPHttpClient.configurations.v1,
  );

  const user = await response.json();

  return {
    id: user.Id,
    title: user.Title,
  };
};

export const getChoiceFieldOptions = async (
  siteUrl: string,
  libraryName: string,
  fieldName: string,
  spHttpClient: SPHttpClient,
): Promise<string[]> => {
  const response = await spHttpClient.get(
    `${siteUrl}/_api/web/lists/getbytitle('${libraryName}')/fields/getbyinternalnameortitle('${fieldName}')`,
    SPHttpClient.configurations.v1,
  );

  const data = await response.json();

  return data.Choices || [];
};
