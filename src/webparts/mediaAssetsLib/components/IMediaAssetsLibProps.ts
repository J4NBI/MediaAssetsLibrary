import { SPHttpClient } from "@microsoft/sp-http";

/**
 * Props für die MediaAssetsLib-Komponente.
 * Enthält die SharePoint-Kontextinformationen und den HTTP-Client für REST-Aufrufe.
 */
export interface IMediaAssetsLibProps {
  /** Kurzbeschreibung des Webparts, die im Property Pane angezeigt wird. */
  description: string;
  /** Gibt an, ob das aktuelle Theme dunkel ist. */
  isDarkTheme: boolean;
  /** Hinweistext zur aktuellen SharePoint- oder Teams-Umgebung. */
  environmentMessage: string;
  /** Gibt an, ob die Komponente im Teams-Kontext gerendert wird. */
  hasTeamsContext: boolean;
  /** Anzeigename des aktuell angemeldeten Benutzers. */
  userDisplayName: string;

  /** SharePoint HTTP-Client für REST- und List-Operationen. */
  spHttpClient: SPHttpClient;
  /** Absolute SharePoint-Site-URL der Zielumgebung. */
  siteUrl: string;
}
