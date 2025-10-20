# easyConnections - Network Connections Monitor (Admin Edition)

**easyConnections** ist ein PowerShell-basiertes Windows-Tool zur Echtzeitüberwachung aktiver Netzwerkverbindungen. Es wurde speziell für Administratoren entwickelt, um schnell und übersichtlich Verbindungen, beteiligte Prozesse und Remoteserver – besonders in Microsoft Exchange zu analysieren.

---

## 🔧 Features

- Anzeige aller aktiven TCP-Verbindungen in Echtzeit
- Filterung nach:
  - Protokoll (LDAP, RPC etc.)
  - Verbindungsstatus (Established, Listening, etc.)
  - Verbindungsrichtung (eingehend/ausgehend)

- Anzeige von:
  - Lokaler und Remote-IP inkl. Ports
  - Remote Hostnamen (z. B. AD-Server, Exchange-Server)
  - Prozessname und PID

- HTML-Export zur Protokollierung
- Paket-Erfassung starten/stoppen
- Anzeige der Prozesse zu den Ports (z. B. w3wp, MSExchangeMailbox, etc.)

---

## 📷 Screenshot

![easyConnections Screenshot](https://github.com/PS-easyIT/easyConnections/blob/main/Screenshot_LDAP_V0.0.2.jpg)

---

## 💻 Systemanforderungen

- Windows 10 / 11 oder Windows Server (2016+)
- PowerShell 5.1 oder höher
