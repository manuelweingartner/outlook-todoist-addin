# Sideload-Anleitung: Mail an Todoist (Dev-Test)

## Voraussetzung: Microsoft 365 Developer-Tenant

Der CMI-Produktiv-Tenant blockiert das Sideloaden eigener Add-ins durch Endbenutzer.
Fur Tests wird ein Microsoft 365 Developer-Tenant benotigt (kostenlos unter
https://developer.microsoft.com/microsoft-365/dev-program).

---

## Lokale Entwicklungsumgebung starten

1. Dev-Zertifikat installieren (einmalig, als Administrator):

   ```bash
   npx office-addin-dev-certs install
   ```

2. Dev-Server starten:

   ```bash
   npm run start
   ```

   Der Webpack-Dev-Server lauft auf `https://localhost:3000`.
   Das Skript versucht automatisch, Outlook zu oeffnen und das Add-in zu laden.

---

## Manuelles Sideloaden im Dev-Tenant (Outlook im Web)

1. Outlook im Web im Developer-Tenant oeffnen: https://outlook.office.com
2. Eine beliebige Mail oeffnen.
3. Oben rechts das 3-Punkte-Menu wahlen > "Apps abrufen" (oder direkt "Apps verwalten").
4. Im Apps-Dialog: "Eigene Add-ins" > "Aus Datei hochladen".
5. Die Datei `manifest.xml` aus dem Projekt-Root auswahlen und bestatigen.
6. Das Add-in erscheint nun in der Toolbar beim Lesen einer Mail als Button "An Todoist-Task".

Hinweis: `manifest.xml` verweist auf `https://localhost:3000`. Der Dev-Server muss
also laufen, wenn das Add-in im Browser verwendet wird.

---

## Troubleshooting

- Zertifikatsfehler im Browser: Dev-Zertifikat neu installieren mit
  `npx office-addin-dev-certs install --machine`.
- Add-in erscheint nicht: Seite neu laden, Cache leeren.
- CORS-Fehler: Todoist erlaubt Browser-Aufrufe mit Bearer-Token direkt (kein Proxy noetig).
