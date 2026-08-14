/**
 * Codifica uma URL de compartilhamento do SharePoint/OneDrive no formato "shareId"
 * exigido pelo Microsoft Graph (`/shares/{shareId}/driveItem`).
 * https://learn.microsoft.com/graph/api/shares-get
 */
export function encodeSharingUrl(url: string): string {
  const base64 = btoa(url);
  const unpadded = base64.replace(/=+$/g, "").replace(/\//g, "_").replace(/\+/g, "-");
  return "u!" + unpadded;
}
